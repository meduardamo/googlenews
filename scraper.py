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
from datetime import datetime, timedelta
from tqdm import tqdm
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
import hashlib
import sys

def setup_driver():
    """Configura e retorna uma instância do WebDriver otimizada"""
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
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

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

def generate_content_hash(title, source, date):
    """Gera um hash único para identificar conteúdo duplicado"""
    content = f"{title.lower().strip()}{source.lower().strip()}{date}"
    return hashlib.md5(content.encode('utf-8')).hexdigest()[:12]

def scrape_news_for_term(driver, termo, periodo_dias=7):
    """Faz scraping das notícias para um termo específico com período configurável"""
    print(f"\n🔍 Buscando notícias para: {termo} (últimos {periodo_dias} dias)")
    
    # Constrói URL com filtro de tempo
    query_text = termo.replace(' ', '+')
    # Adiciona filtro temporal para melhor relevância
    when_param = "when:7d" if periodo_dias <= 7 else "when:1m"
    link = f"https://news.google.com/search?q={query_text}+{when_param}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"
    
    try:
        driver.get(link)
        time.sleep(3)

        print("📜 Fazendo scroll da página...")
        scroll_attempts = 0
        max_attempts = 8  # Reduzido para ser mais eficiente
        last_height = driver.execute_script("return document.body.scrollHeight")
        
        while scroll_attempts < max_attempts:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1.5)  # Tempo otimizado
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
            scroll_attempts += 1

        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        news_items = soup.select('div.UW0SDc, article')

        noticias = []
        noticias_hash_set = set()  # Para evitar duplicatas durante o scraping
        root = 'https://news.google.com'
        cutoff_date = datetime.now() - timedelta(days=periodo_dias)

        for item in news_items:
            try:
                title = item.find('a', class_='JtKRv') or item.find('h3') or item.find('h4')
                link_item = item.find("a", href=True)
                publisher = item.find('div', class_='vr1PYe') or item.find('div', class_='wsLqz')
                data_publicacao_tag = item.find('time', class_='hvbAAd') or item.find('time')

                if not title or not link_item:
                    continue

                # Processamento de data aprimorado
                datetime_string = data_publicacao_tag['datetime'] if data_publicacao_tag and data_publicacao_tag.get('datetime') else None
                data_publicacao = None
                data_dt = None
                
                if datetime_string:
                    try:
                        data_dt = datetime.strptime(datetime_string, '%Y-%m-%dT%H:%M:%SZ')
                        # Filtra notícias muito antigas no próprio scraping
                        if data_dt < cutoff_date:
                            continue
                        data_publicacao = data_dt.strftime('%d/%m/%Y %H:%M')
                    except ValueError:
                        # Tenta outros formatos de data
                        try:
                            data_dt = datetime.strptime(datetime_string.split('T')[0], '%Y-%m-%d')
                            data_publicacao = data_dt.strftime('%d/%m/%Y')
                        except ValueError:
                            continue  # Skip se não conseguir parsear a data

                title_text = title.text.strip()
                source_text = publisher.text.strip() if publisher else 'Fonte não identificada'
                
                # Gera hash para detectar duplicatas
                content_hash = generate_content_hash(title_text, source_text, data_publicacao or '')
                if content_hash in noticias_hash_set:
                    continue
                noticias_hash_set.add(content_hash)

                # Melhora o link (remove parâmetros desnecessários)
                link_href = link_item['href']
                if link_href.startswith('./'):
                    clean_link = root + link_href[1:]
                elif link_href.startswith('/'):
                    clean_link = root + link_href
                else:
                    clean_link = link_href

                noticia = {
                    'Título': title_text,
                    'Fonte': source_text,
                    'Data de Publicação': data_publicacao or 'Data não encontrada',
                    'Link': clean_link,
                    'Termo de Busca': termo,
                    'Hash de Conteúdo': content_hash,
                    'Data de Coleta': datetime.now().strftime('%d/%m/%Y %H:%M'),
                    'Relevância': 'Alta' if any(word in title_text.lower() for word in termo.lower().split()) else 'Média'
                }
                noticias.append(noticia)
                
            except Exception as e:
                print(f"⚠️ Erro ao processar notícia: {e}")
                continue

        print(f"📰 {len(noticias)} notícias encontradas para '{termo}'")
        return noticias
        
    except Exception as e:
        print(f"❌ Erro ao fazer scraping para '{termo}': {e}")
        return []

def remove_duplicates_advanced(df):
    """Remove duplicatas usando múltiplos critérios"""
    if df.empty:
        return df
    
    # Remove duplicatas por hash primeiro
    df_unique = df.drop_duplicates(subset=['Hash de Conteúdo']).copy()
    
    # Remove duplicatas por similaridade de título (90% similar)
    df_final = []
    titles_processed = set()
    
    for _, row in df_unique.iterrows():
        title = row['Título'].lower().strip()
        is_duplicate = False
        
        for processed_title in titles_processed:
            # Verifica similaridade simples (pode ser melhorado com algoritmos mais sofisticados)
            if len(set(title.split()) & set(processed_title.split())) / len(set(title.split()) | set(processed_title.split())) > 0.8:
                is_duplicate = True
                break
        
        if not is_duplicate:
            titles_processed.add(title)
            df_final.append(row)
    
    return pd.DataFrame(df_final)

def _gspread_client_from_env():
    """Cria client do gspread com autenticação otimizada"""
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

def setup_google_sheets(gc, spreadsheet_key):
    """Configura as abas do Google Sheets de forma padronizada"""
    try:
        spreadsheet = gc.open_by_key(spreadsheet_key)
        
        # Configuração das abas
        sheets_config = {
            'Notícias Recentes': {'rows': 1000, 'cols': 10},
            'Dashboard': {'rows': 50, 'cols': 6},
            'Histórico': {'rows': 5000, 'cols': 10},
            'Configurações': {'rows': 20, 'cols': 4}
        }
        
        existing_sheets = [ws.title for ws in spreadsheet.worksheets()]
        
        for sheet_name, config in sheets_config.items():
            if sheet_name not in existing_sheets:
                print(f"📄 Criando aba: {sheet_name}")
                spreadsheet.add_worksheet(title=sheet_name, rows=config['rows'], cols=config['cols'])
        
        return spreadsheet
        
    except Exception as e:
        print(f"❌ Erro ao configurar Google Sheets: {e}")
        raise

def upload_to_google_sheets(df_geral_final, resumo_coletas):
    """Upload otimizado para Google Sheets com múltiplas abas"""
    try:
        print("📤 Enviando dados para Google Sheets...")
        gc = _gspread_client_from_env()
        
        SPREADSHEET_KEY = '1G81BndSPpnViMDxRKQCth8PwK0xmAwH-w-T7FjgnwcY'
        spreadsheet = setup_google_sheets(gc, SPREADSHEET_KEY)
        
        # Aba principal - Notícias Recentes
        ws_noticias = spreadsheet.worksheet('Notícias Recentes')
        
        if not df_geral_final.empty:
            # Ordena por data de publicação (mais recentes primeiro)
            df_sorted = df_geral_final.sort_values('Data de Publicação', ascending=False)
            set_with_dataframe(ws_noticias, df_sorted)
        else:
            # Cabeçalhos vazios se não houver dados
            empty_df = pd.DataFrame(columns=['Título', 'Fonte', 'Data de Publicação', 'Link', 'Termo de Busca', 'Data de Coleta', 'Relevância'])
            set_with_dataframe(ws_noticias, empty_df)
        
        # Aba Dashboard - Resumo executivo
        ws_dashboard = spreadsheet.worksheet('Dashboard')
        
        # Cria métricas para o dashboard
        total_noticias = len(df_geral_final) if not df_geral_final.empty else 0
        dashboard_data = [
            ['Métrica', 'Valor', 'Última Atualização'],
            ['Total de Notícias Coletadas', total_noticias, datetime.now().strftime('%d/%m/%Y %H:%M')],
            ['Termos Monitorados', len(resumo_coletas), ''],
            ['Status', 'Ativo', ''],
        ]
        
        # Adiciona resumo por termo
        dashboard_data.append(['', '', ''])
        dashboard_data.append(['Resumo por Termo', '', ''])
        
        for item in resumo_coletas:
            dashboard_data.append([item['Termo de Busca'], f"{item['Notícias Coletadas']} notícias", ''])
        
        ws_dashboard.clear()
        for i, row in enumerate(dashboard_data, 1):
            ws_dashboard.insert_row(row, i)
        
        print("✅ Dados enviados para Google Sheets com sucesso")
        print(f"🔗 Acesse: https://docs.google.com/spreadsheets/d/{SPREADSHEET_KEY}")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao enviar para Google Sheets: {e}")
        return False

def main():
    """Função principal otimizada"""
    print("🚀 Iniciando scraper de notícias otimizado...")
    print(f"⏰ Execução iniciada em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    driver = setup_driver()

    # Configuração flexível de termos de busca
    search_terms = [
        'PNE',
        'Plano Nacional de Educação', 
        'Saúde Mental',
        # Adicione novos termos aqui conforme necessário
    ]
    
    resultados_por_termo = []
    resumo_coletas = []
    periodo_dias = 7  # Configurável

    try:
        print(f"🎯 Monitorando {len(search_terms)} termos nas últimas {periodo_dias} dias")
        
        for termo in tqdm(search_terms, desc="🔎 Processando termos"):
            noticias = scrape_news_for_term(driver, termo, periodo_dias)
            
            if noticias:
                df_noticias = pd.DataFrame(noticias)
                resultados_por_termo.append(df_noticias)
                print(f"✅ {len(df_noticias)} notícias coletadas para '{termo}'")
            else:
                print(f"⚠️ Nenhuma notícia encontrada para '{termo}'")
            
            resumo_coletas.append({
                'Termo de Busca': termo, 
                'Notícias Coletadas': len(noticias),
                'Status': 'Sucesso' if noticias else 'Sem resultados'
            })
            
            # Pequena pausa entre termos para evitar rate limiting
            time.sleep(2)

        # Consolida e limpa os dados
        if resultados_por_termo:
            df_geral = pd.concat(resultados_por_termo, ignore_index=True)
            print(f"🔄 Removendo duplicatas de {len(df_geral)} notícias...")
            df_geral_final = remove_duplicates_advanced(df_geral)
            print(f"✅ {len(df_geral_final)} notícias únicas após limpeza")
        else:
            df_geral_final = pd.DataFrame()

        # Upload direto para Google Sheets (sem Excel local)
        sheets_success = upload_to_google_sheets(df_geral_final, resumo_coletas)

        # Relatório final
        total_noticias = len(df_geral_final) if not df_geral_final.empty else 0
        print(f"\n📊 RESUMO DA EXECUÇÃO:")
        print(f"📰 Total de notícias únicas: {total_noticias}")
        print(f"🗓️ Período analisado: últimos {periodo_dias} dias")
        print(f"📤 Google Sheets atualizado: {'✅' if sheets_success else '❌'}")
        print(f"⏱️ Execução finalizada em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

        if total_noticias == 0:
            print("⚠️ Nenhuma notícia relevante encontrada no período")
        else:
            print(f"🎉 Monitoramento concluído com sucesso!")
            
        # Exit code baseado no sucesso
        sys.exit(0 if sheets_success else 1)
        
    except Exception as e:
        print(f"❌ Erro crítico na execução: {e}")
        sys.exit(1)
    finally:
        driver.quit()
        print("🔚 WebDriver encerrado")

if __name__ == "__main__":
    main()
