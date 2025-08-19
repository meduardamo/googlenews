# -*- coding: utf-8 -*-
import os
import json
import time
import shutil
import hashlib
from zoneinfo import ZoneInfo
import pandas as pd
import requests

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from datetime import datetime, timezone, timedelta
from tqdm import tqdm
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials  # <— moderno
import sys

# =========================
# Utilidades de horário
# =========================
TZ_BR = ZoneInfo("America/Sao_Paulo")

def agora_brasilia():
    return datetime.now(tz=TZ_BR)

def fmt_brasilia(dt: datetime) -> str:
    return dt.astimezone(TZ_BR).strftime("%d/%m/%Y %H:%M")

# =========================
# Selenium / Driver
# =========================
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
# ↓ não precisa mais do Service / webdriver_manager

def setup_driver():
    """Configura e retorna uma instância do WebDriver usando Selenium Manager (sem webdriver_manager)."""
    print("🔧 Configurando WebDriver (Selenium Manager)...")
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

    # Em Actions, essas flags evitam throttling quando headless
    if os.getenv('GITHUB_ACTIONS'):
        options.add_argument("--disable-background-timer-throttling")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--disable-backgrounding-occluded-windows")

    # Se o setup-chrome já pôs o Chrome no PATH, o Selenium Manager resolve o driver certo sozinho
    try:
        driver = webdriver.Chrome(options=options)
        print("✅ WebDriver configurado com sucesso (Selenium Manager)")
        return driver
    except Exception as e:
        print(f"❌ Erro ao configurar WebDriver: {e}")
        sys.exit(1)

# =========================
# Fingerprints no Google Sheets
# =========================
def _gspread_client_from_env():
    """Lê credenciais do env (aceita GCP_SERVICE_ACCOUNT_JSON ou GOOGLE_APPLICATION_CREDENTIALS_JSON)"""
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

SPREADSHEET_KEY = '1G81BndSPpnViMDxRKQCth8PwK0xmAwH-w-T7FjgnwcY'
WORKSHEET_DATA = 'google notícias'
WORKSHEET_FP = '_fingerprints'  # aba "oculta" lógica (não precisa aparecer para o usuário)

def _ensure_worksheets(spreadsheet):
    try:
        ws_data = spreadsheet.worksheet(WORKSHEET_DATA)
    except gspread.WorksheetNotFound:
        ws_data = spreadsheet.add_worksheet(title=WORKSHEET_DATA, rows="100", cols="20")
    try:
        ws_fp = spreadsheet.worksheet(WORKSHEET_FP)
    except gspread.WorksheetNotFound:
        ws_fp = spreadsheet.add_worksheet(title=WORKSHEET_FP, rows="100", cols="2")
        ws_fp.update('A1:B1', [['fp', 'created_at_brt']])
    return ws_data, ws_fp

def load_fingerprints(gc) -> set:
    """Lê os fingerprints existentes da aba _fingerprints e retorna um set."""
    print("🧩 Carregando fingerprints do Google Sheets...")
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    _, ws_fp = _ensure_worksheets(spreadsheet)
    vals = ws_fp.get_all_records()
    fps = {row.get('fp') for row in vals if row.get('fp')}
    print(f"🧩 Fingerprints carregados: {len(fps)}")
    return fps

def append_fingerprints(gc, new_fps: list[str]):
    if not new_fps:
        return
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    _, ws_fp = _ensure_worksheets(spreadsheet)
    now_brt = fmt_brasilia(agora_brasilia())
    rows = [[fp, now_brt] for fp in new_fps]
    ws_fp.append_rows(rows, value_input_option='RAW')
    print(f"🧩 Fingerprints adicionados: {len(new_fps)}")

# =========================
# Resolvedor de URL final
# =========================
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0 Safari/537.36"
})

def resolve_final_url(driver, url: str) -> str:
    """Abre o link do Google News no navegador e pega a URL final redirecionada."""
    try:
        driver.get(url)
        time.sleep(2)  # dá tempo pro redirect acontecer
        return driver.current_url
    except Exception:
        return url

# =========================
# Scraper
# =========================
def scrape_news_for_term(driver, termo):
    """Faz scraping das notícias para um termo específico"""
    print(f"\n🔍 Buscando notícias para: {termo}")
    query_text = termo.replace(' ', '+')
    link = f"https://news.google.com/search?q={query_text}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"
    try:
        driver.get(link)
        time.sleep(3)

        # scroll para carregar
        print("📜 Fazendo scroll da página...")
        for _ in range(10):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        news_items = soup.select('div.UW0SDc, article')

        noticias = []
        root = 'https://news.google.com'
        for item in news_items:
            try:
                title = item.find('a', class_='JtKRv') or item.find('h3') or item.find('h4')
                link_item = item.find("a", href=True)
                publisher = item.find('div', class_='vr1PYe') or item.find('div', class_='wsLqz')
                time_tag = item.find('time', class_='hvbAAd') or item.find('time')

                # datetime ISO em UTC se disponível (ex.: 2025-08-19T12:34:56Z)
                dt_utc = None
                if time_tag and time_tag.get('datetime'):
                    try:
                        dt_utc = datetime.strptime(time_tag['datetime'], '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc)
                    except Exception:
                        dt_utc = None

                # link bruto do Google News
                link_bruto = root + link_item['href'][1:] if link_item and link_item.get('href') else None
                if not link_bruto:
                    continue

                # URL final resolvida
                url_final = resolve_final_url(driver, link_bruto)

                noticia = {
                    'Título': title.text.strip() if title else 'Título não encontrado',
                    'Fonte': publisher.text.strip() if publisher else 'Fonte não encontrada',
                    # mantemos data de publicação como dd/mm/aaaa; usamos dt_utc para filtro 24h
                    'Data de Publicação': dt_utc.astimezone(TZ_BR).strftime('%d/%m/%Y') if dt_utc else 'Data não encontrada',
                    'Link': link_bruto,
                    'URL Final': url_final,
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

def filter_recent_24h(df_noticias):
    """Mantém apenas notícias cujo timestamp do Google News está nas últimas 24h."""
    if df_noticias.empty:
        return pd.DataFrame()

    # Tentar recuperar a coluna ISO original via 'Link'? Não vem. Então refazemos parse do 'Data de Publicação' não ajuda.
    # Solução: durante o scrape acima, guardamos apenas a data formatada. Para o filtro 24h real, vamos re-scrapear as <time>?
    # Melhor: salvar 'Publicado_UTC' internamente durante o scrape.
    # Ajuste: se não existir, tudo passa (fallback seguro).

    if 'Publicado_UTC' not in df_noticias.columns:
        # compatibilidade: cria vazia
        df_noticias['Publicado_UTC'] = pd.NaT

    now_utc = datetime.now(timezone.utc)
    def _in_24h(val):
        if pd.isna(val):
            return True  # se não temos o horário, não bloqueia (ou mude para False se preferir ser estrito)
        return (now_utc - val) <= timedelta(hours=24)

    mask = df_noticias['Publicado_UTC'].apply(_in_24h)
    return df_noticias[mask].copy()

# Pequena correção: vamos ajustar o scraper para também preencher 'Publicado_UTC'
def scrape_news_for_term(driver, termo):
    print(f"\n🔍 Buscando notícias para: {termo}")
    query_text = termo.replace(' ', '+')
    link = f"https://news.google.com/search?q={query_text}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"
    try:
        driver.get(link)
        time.sleep(3)
        print("📜 Fazendo scroll da página...")
        for _ in range(10):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        news_items = soup.select('div.UW0SDc, article')

        noticias = []
        root = 'https://news.google.com'
        for item in news_items:
            try:
                title = item.find('a', class_='JtKRv') or item.find('h3') or item.find('h4')
                link_item = item.find("a", href=True)
                publisher = item.find('div', class_='vr1PYe') or item.find('div', class_='wsLqz')
                time_tag = item.find('time', class_='hvbAAd') or item.find('time')

                dt_utc = None
                if time_tag and time_tag.get('datetime'):
                    try:
                        dt_utc = datetime.strptime(time_tag['datetime'], '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc)
                    except Exception:
                        dt_utc = None

                link_bruto = root + link_item['href'][1:] if link_item and link_item.get('href') else None
                if not link_bruto:
                    continue

                url_final = resolve_final_url(link_bruto)

                noticias.append({
                    'Título': title.text.strip() if title else 'Título não encontrado',
                    'Fonte': publisher.text.strip() if publisher else 'Fonte não encontrada',
                    'Data de Publicação': dt_utc.astimezone(TZ_BR).strftime('%d/%m/%Y') if dt_utc else 'Data não encontrada',
                    'Link': link_bruto,
                    'URL Final': url_final,
                    'Termo de Busca': termo,
                    'Publicado_UTC': dt_utc
                })
            except Exception as e:
                print(f"⚠️ Erro ao processar notícia: {e}")
                continue
        return noticias
    except Exception as e:
        print(f"❌ Erro ao fazer scraping para '{termo}': {e}")
        return []

def make_fingerprint(row) -> str:
    """
    Gera um fingerprint estável da notícia.
    Usamos URL Final (ou Link) + Título, normalizados.
    """
    base = (row.get('URL Final') or row.get('Link') or '') + '|' + (row.get('Título') or '')
    base = base.strip().lower()
    return hashlib.sha256(base.encode('utf-8')).hexdigest()

# =========================
# Persistência local (Excel) e Sheets
# =========================
def save_to_excel(df_geral_final, resumo_coletas, filename='noticias_PNE_Planos_Educacao.xlsx'):
    """Salva os dados em Excel"""
    try:
        with pd.ExcelWriter(filename) as writer:
            (df_geral_final if not df_geral_final.empty else pd.DataFrame(
                columns=['Título','Fonte','Data de Publicação','Link','URL Final','Termo de Busca','Coletado em'])) \
                .to_excel(writer, sheet_name='Noticias', index=False)
            pd.DataFrame(resumo_coletas).to_excel(writer, sheet_name='Resumo', index=False)
        print(f"✅ Arquivo Excel '{filename}' salvo com sucesso")
        return True
    except Exception as e:
        print(f"❌ Erro ao salvar Excel: {e}")
        return False

def upload_to_google_sheets(df_geral_final):
    """Upload para Google Sheets (google-auth)"""
    try:
        print("📤 Enviando dados para Google Sheets...")
        gc = _gspread_client_from_env()
        spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
        ws_data, _ = _ensure_worksheets(spreadsheet)
        set_with_dataframe(
            ws_data,
            df_geral_final if not df_geral_final.empty else pd.DataFrame(
                columns=['Título','Fonte','Data de Publicação','Link','URL Final','Termo de Busca','Coletado em'])
        )
        print("✅ Dados enviados para Google Sheets com sucesso")
        return True, gc
    except Exception as e:
        print(f"❌ Erro ao enviar para Google Sheets: {e}")
        return False, None

# =========================
# Main
# =========================
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
            df_noticias_24h = filter_recent_24h(df_noticias)
            resultados_por_termo.append(df_noticias_24h)
            resumo_coletas.append({'Termo de Busca': termo, 'Notícias Coletadas (24h)': len(df_noticias_24h)})

        df_geral = pd.concat(resultados_por_termo, ignore_index=True) if resultados_por_termo else pd.DataFrame()

        # dedupe intra-execução (preferindo URL Final)
        if not df_geral.empty:
            if 'URL Final' in df_geral.columns:
                df_geral = df_geral.drop_duplicates(subset=['URL Final']).copy()
            else:
                df_geral = df_geral.drop_duplicates(subset=['Link']).copy()

        # adiciona coluna "Coletado em" (Brasília)
        coletado_em_str = fmt_brasilia(agora_brasilia())
        if not df_geral.empty:
            df_geral['Coletado em'] = coletado_em_str

        # remove colunas internas
        if 'Publicado_UTC' in df_geral.columns:
            df_geral_final = df_geral.drop(columns=['Publicado_UTC'])
        else:
            df_geral_final = df_geral

        # ====== Deduplicação entre execuções via fingerprints no Sheets ======
        new_fps = []
        if not df_geral_final.empty:
            df_geral_final['__fp'] = df_geral_final.apply(make_fingerprint, axis=1)
            try:
                gc = _gspread_client_from_env()
                existing = load_fingerprints(gc)
                mask_novos = ~df_geral_final['__fp'].isin(existing)
                df_geral_final = df_geral_final[mask_novos].copy()
                new_fps = df_geral_final['__fp'].tolist()
            except Exception as e:
                print(f"⚠️ Não foi possível usar fingerprints remotos (seguindo sem): {e}")

            # limpa a coluna técnica
            if '__fp' in df_geral_final.columns:
                df_geral_final.drop(columns=['__fp'], inplace=True)

        # salvar Excel
        excel_saved = save_to_excel(df_geral_final, resumo_coletas)

        # enviar para Sheets
        sheets_uploaded, gc_for_fp = upload_to_google_sheets(df_geral_final)

        # se conseguimos enviar e temos novos fingerprints, anexa na aba _fingerprints
        if sheets_uploaded and new_fps:
            try:
                append_fingerprints(gc_for_fp or _gspread_client_from_env(), new_fps)
            except Exception as e:
                print(f"⚠️ Falha ao registrar fingerprints: {e}")

        total_noticias = len(df_geral_final) if not df_geral_final.empty else 0
        print(f"\n📊 RESUMO DA EXECUÇÃO:")
        print(f"📰 Total de notícias novas (24h, sem repetição): {total_noticias}")
        print(f"🕒 Coletado em (BRT): {coletado_em_str}")
        print(f"📁 Excel salvo: {'✅' if excel_saved else '❌'}")
        print(f"📤 Google Sheets atualizado: {'✅' if sheets_uploaded else '❌'}")
        if total_noticias == 0:
            print("⚠️ Nenhuma notícia nova encontrada nas últimas 24h (após deduplicação).")
            sys.exit(0)
    except Exception as e:
        print(f"❌ Erro geral na execução: {e}")
        sys.exit(1)
    finally:
        driver.quit()
        print("🏁 Scraper finalizado")

if __name__ == "__main__":
    main()
