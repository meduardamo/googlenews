# -*- coding: utf-8 -*-
import os
import json
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from datetime import datetime
from tqdm import tqdm
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials  # <— moderno
import sys

def setup_driver():
    """Configura e retorna uma instância do WebDriver"""
    print("🔧 Configurando WebDriver...")
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920x1080")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-plugins")
    options.add_argument("--disable-images")
    options.add_argument("--disable-javascript")

    if os.getenv('GITHUB_ACTIONS'):
        options.add_argument("--disable-background-timer-throttling")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--disable-backgrounding-occluded-windows")

    chrome_path = shutil.which("google-chrome") or shutil.which("chrome") or shutil.which("chromium")
    if chrome_path:
        options.binary_location = chrome_path
        print(f"✅ Chrome encontrado em: {chrome_path}")

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        print("✅ WebDriver configurado com sucesso")
        return driver
    except Exception as e:
        print(f"❌ Erro ao configurar WebDriver: {e}")
        sys.exit(1)

def scrape_news_for_term(driver, termo):
    """Faz scraping das notícias para um termo específico"""
    print(f"\n🔍 Buscando notícias para: {termo}")
    query_text = termo.replace(' ', '+')
    link = f"https://news.google.com/search?q={query_text}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"
    try:
        driver.get(link)
        time.sleep(3)

        print("📜 Fazendo scroll da página...")
        scroll_attempts = 0
        max_attempts = 10
        last_height = driver.execute_script("return document.body.scrollHeight")
        while scroll_attempts < max_attempts:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                print("📜 Final do conteúdo alcançado.")
                break
            last_height = new_height
            scroll_attempts += 1

        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        news_items = soup.select('div.UW0SDc, article')

        noticias = []
        root = 'https://news.google.com'

        for item in news_items:
            try:
                title = item.find('a', class_='JtKRv') or item.find('h3') or item.find('h4')
                link_item = item.find("a", href=True)
                publisher = item.find('div', class_='vr1PYe') or item.find('div', class_='wsLqz')
                data_publicacao_tag = item.find('time', class_='hvbAAd') or item.find('time')

                datetime_string = data_publicacao_tag['datetime'] if data_publicacao_tag and data_publicacao_tag.get('datetime') else None
                data_publicacao = None
                if datetime_string:
                    try:
                        dt = datetime.strptime(datetime_string, '%Y-%m-%dT%H:%M:%SZ')
                        data_publicacao = dt.strftime('%d/%m/%Y')
                    except ValueError:
                        data_publicacao = 'Data inválida'

                noticia = {
                    'Título': title.text.strip() if title else 'Título não encontrado',
                    'Fonte': publisher.text.strip() if publisher else 'Fonte não encontrada',
                    'Data de Publicação': data_publicacao if data_publicacao else 'Data não encontrada',
                    'Link': root + link_item['href'][1:] if link_item and link_item.get('href') else 'Link não encontrado',
                    'Termo de Busca': termo
                }
                noticias.append(noticia)
            except Exception as e:
                print(f"⚠️ Erro ao processar notícia: {e}")
                continue

        return noticias
    except Exception as e:
        print(f"❌ Erro ao fazer scraping para '{termo}': {e}")
        return []

def filter_recent_news(df_noticias):
    """Filtra notícias das últimas 24h"""
    if df_noticias.empty:
        return pd.DataFrame()
    hoje = pd.Timestamp.now()
    df_noticias['Data Convertida'] = pd.to_datetime(df_noticias['Data de Publicação'], format='%d/%m/%Y', errors='coerce')
    df_noticias['Dias de Diferença'] = (hoje - df_noticias['Data Convertida']).dt.days
    return df_noticias[df_noticias['Dias de Diferença'] <= 1].copy()

def save_to_excel(df_geral_final, resumo_coletas, filename='noticias_PNE_Planos_Educacao.xlsx'):
    """Salva os dados em Excel"""
    try:
        with pd.ExcelWriter(filename) as writer:
            (df_geral_final if not df_geral_final.empty else 
             pd.DataFrame(columns=['Título','Fonte','Data de Publicação','Link','Termo de Busca']))\
                .to_excel(writer, sheet_name='Noticias', index=False)
            pd.DataFrame(resumo_coletas).to_excel(writer, sheet_name='Resumo', index=False)
        print(f"✅ Arquivo Excel '{filename}' salvo com sucesso")
        return True
    except Exception as e:
        print(f"❌ Erro ao salvar Excel: {e}")
        return False

def _gspread_client_from_env():
    """
    Lê credenciais do env (aceita GCP_SERVICE_ACCOUNT_JSON ou GOOGLE_APPLICATION_CREDENTIALS_JSON),
    normaliza \\n -> \n e cria o client do gspread com google-auth.
    """
    raw = os.environ.get("GCP_SERVICE_ACCOUNT_JSON") or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if not raw:
        raise RuntimeError("Secret GCP_SERVICE_ACCOUNT_JSON ou GOOGLE_APPLICATION_CREDENTIALS_JSON não encontrado.")
    info = json.loads(raw)
    if "private_key" in info and "\\n" in info["private_key"]:
        info["private_key"] = info["private_key"].replace("\\n", "\n")
    info["token_uri"] = "https://oauth2.googleapis.com/token"
    creds = Credentials.from_service_account_info(info, scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    return gspread.authorize(creds)

def upload_to_google_sheets(df_geral_final):
    """Upload para Google Sheets (google-auth)"""
    try:
        print("📤 Enviando dados para Google Sheets...")
        gc = _gspread_client_from_env()

        SPREADSHEET_KEY = '1G81BndSPpnViMDxRKQCth8PwK0xmAwH-w-T7FjgnwcY'
        worksheet_name = 'google notícias'

        spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="100", cols="20")

        set_with_dataframe(
            worksheet,
            df_geral_final if not df_geral_final.empty else 
            pd.DataFrame(columns=['Título','Fonte','Data de Publicação','Link','Termo de Busca'])
        )
        print("✅ Dados enviados para Google Sheets com sucesso")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar para Google Sheets: {e}")
        return False

def main():
    """Função principal"""
    print("🚀 Iniciando scraper de notícias...")
    driver = setup_driver()

    search_terms = ['PNE', 'Plano Nacional de Educação', 'Saúde Mental']
    resultados_por_termo = []
    resumo_coletas = []

    try:
        for termo in tqdm(search_terms, desc="🔎 Buscando termos"):
            noticias = scrape_news_for_term(driver, termo)
            df_noticias = pd.DataFrame(noticias)
            df_noticias_24h = filter_recent_news(df_noticias)
            print(f"✅ {len(df_noticias_24h)} notícias coletadas para o termo '{termo}'")
            resultados_por_termo.append(df_noticias_24h)
            resumo_coletas.append({'Termo de Busca': termo, 'Notícias Coletadas': len(df_noticias_24h)})

        df_geral = pd.concat(resultados_por_termo, ignore_index=True) if resultados_por_termo else pd.DataFrame()
        if not df_geral.empty and 'Link' in df_geral.columns:
            df_geral = df_geral.drop_duplicates(subset='Link').copy()

        cols_drop = [c for c in ['Data Convertida', 'Dias de Diferença'] if c in df_geral.columns]
        df_geral_final = df_geral.drop(columns=cols_drop) if not df_geral.empty else pd.DataFrame()

        excel_saved = save_to_excel(df_geral_final, resumo_coletas)
        sheets_uploaded = upload_to_google_sheets(df_geral_final)

        total_noticias = len(df_geral_final) if not df_geral_final.empty else 0
        print(f"\n📊 RESUMO DA EXECUÇÃO:")
        print(f"📰 Total de notícias coletadas: {total_noticias}")
        print(f"📁 Excel salvo: {'✅' if excel_saved else '❌'}")
        print(f"📤 Google Sheets atualizado: {'✅' if sheets_uploaded else '❌'}")

        if total_noticias == 0:
            print("⚠️ Nenhuma notícia encontrada nas últimas 24h")
            sys.exit(0)
    except Exception as e:
        print(f"❌ Erro geral na execução: {e}")
        sys.exit(1)
    finally:
        driver.quit()
        print("🏁 Scraper finalizado")

if __name__ == "__main__":
    main()
