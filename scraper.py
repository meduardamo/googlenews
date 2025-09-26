# -*- coding: utf-8 -*-
import os
import json
import time
import hashlib
from zoneinfo import ZoneInfo
import pandas as pd
import requests

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from datetime import datetime, timezone, timedelta
from tqdm import tqdm
import gspread
from google.oauth2.service_account import Credentials
import sys

# =========================
# Config e Constantes
# =========================
TZ_BR = ZoneInfo("America/Sao_Paulo")

# Agora a chave vem de uma variável de ambiente (secret)
SPREADSHEET_KEY = os.environ.get('PLANILHA')
if not SPREADSHEET_KEY:
    raise RuntimeError("Secret PLANILHA não encontrado. Configure o secret no GitHub Actions.")

WORKSHEET_DATA = 'google notícias'
WORKSHEET_FP = '_fingerprints'  # armazena fp e timestamp

HEADERS_DATA = ['Título','Fonte','Data de Publicação','Link','URL Final','Termo de Busca','Coletado em']
HEADERS_FP   = ['fp','created_at_brt']

SEARCH_TERMS = ['PNE', 'Plano Nacional de Educação', 'Saúde Mental']

# =========================
# Utilidades de horário
# =========================
def agora_brasilia():
    return datetime.now(tz=TZ_BR)

def fmt_brasilia(dt: datetime) -> str:
    return dt.astimezone(TZ_BR).strftime("%d/%m/%Y %H:%M")

# =========================
# Selenium / Driver
# =========================
def setup_driver():
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
    if os.getenv('GITHUB_ACTIONS'):
        options.add_argument("--disable-background-timer-throttling")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--disable-backgrounding-occluded-windows")
    try:
        driver = webdriver.Chrome(options=options)
        print("✅ WebDriver configurado com sucesso (Selenium Manager)")
        return driver
    except Exception as e:
        print(f"❌ Erro ao configurar WebDriver: {e}")
        sys.exit(1)

# =========================
# Google Sheets helpers
# =========================
def _gspread_client_from_env():
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

def _col_letter(n_cols: int) -> str:
    # como temos poucas colunas, um conversor simples (1->A, 7->G)
    return chr(64 + n_cols)

def _ensure_worksheets(spreadsheet):
    # DATA
    try:
        ws_data = spreadsheet.worksheet(WORKSHEET_DATA)
    except gspread.WorksheetNotFound:
        ws_data = spreadsheet.add_worksheet(title=WORKSHEET_DATA, rows="100", cols=str(len(HEADERS_DATA)))
        ws_data.update(f'A1:{_col_letter(len(HEADERS_DATA))}1', [HEADERS_DATA])  # cabeçalho
    # garante cabeçalho mesmo se a aba já existia sem header correto
    first_row = ws_data.row_values(1)
    if [c.strip().lower() for c in first_row] != [c.lower() for c in HEADERS_DATA]:
        ws_data.update(f'A1:{_col_letter(len(HEADERS_DATA))}1', [HEADERS_DATA])

    # FINGERPRINTS
    try:
        ws_fp = spreadsheet.worksheet(WORKSHEET_FP)
    except gspread.WorksheetNotFound:
        ws_fp = spreadsheet.add_worksheet(title=WORKSHEET_FP, rows="100", cols=str(len(HEADERS_FP)))
        ws_fp.update(f'A1:{_col_letter(len(HEADERS_FP))}1', [HEADERS_FP])        # cabeçalho
    first_row_fp = ws_fp.row_values(1)
    if [c.strip().lower() for c in first_row_fp] != [c.lower() for c in HEADERS_FP]:
        ws_fp.update(f'A1:{_col_letter(len(HEADERS_FP))}1', [HEADERS_FP])

    return ws_data, ws_fp

def load_fingerprints(gc) -> set:
    print("🧩 Carregando fingerprints do Google Sheets...")
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    _, ws_fp = _ensure_worksheets(spreadsheet)
    vals = ws_fp.get_all_records()
    fps = {row.get('fp') for row in vals if row.get('fp')}
    print(f"🧩 Fingerprints carregados: {len(fps)}")
    return fps

def insert_fingerprints(gc, new_fps: list[str]):
    if not new_fps:
        return
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    _, ws_fp = _ensure_worksheets(spreadsheet)
    now_brt = fmt_brasilia(agora_brasilia())
    rows = [[fp, now_brt] for fp in new_fps]
    # Insere fingerprints logo após o cabeçalho
    ws_fp.insert_rows(rows, row=2, value_input_option='RAW')
    print(f"🧩 Fingerprints inseridos: {len(new_fps)}")

def insert_news_rows(gc, df_novos: pd.DataFrame):
    """Insert na aba 'google notícias' logo após o cabeçalho."""
    if df_novos.empty:
        print("ℹ️ Nada novo para adicionar na aba de dados.")
        return
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    ws_data, _ = _ensure_worksheets(spreadsheet)

    # normaliza/ordena colunas
    for col in HEADERS_DATA:
        if col not in df_novos.columns:
            df_novos[col] = ''
    df_novos = df_novos[HEADERS_DATA].fillna('')

    # Insere linhas a partir da linha 2 (logo após o cabeçalho)
    ws_data.insert_rows(df_novos.values.tolist(), row=2, value_input_option='RAW')
    print(f"✅ Linhas inseridas no início da aba '{WORKSHEET_DATA}': {len(df_novos)}")

# =========================
# Resolvedor de URL final
# =========================
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0 Safari/537.36"
})

def resolve_final_url(url: str) -> str:
    try:
        resp = SESSION.get(url, allow_redirects=True, timeout=12)
        return resp.url
    except Exception:
        return url

# =========================
# Scraper
# =========================
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

def filter_recent_24h(df_noticias):
    if df_noticias.empty:
        return pd.DataFrame()
    if 'Publicado_UTC' not in df_noticias.columns:
        df_noticias['Publicado_UTC'] = pd.NaT

    now_utc = datetime.now(timezone.utc)
    def _in_24h(val):
        if pd.isna(val):
            # se não tem timestamp, mantemos (Google News às vezes falha)
            return True
        return (now_utc - val) <= timedelta(hours=24)

    mask = df_noticias['Publicado_UTC'].apply(_in_24h)
    return df_noticias[mask].copy()

def make_fingerprint(row) -> str:
    base = (row.get('URL Final') or row.get('Link') or '') + '|' + (row.get('Título') or '')
    base = base.strip().lower()
    return hashlib.sha256(base.encode('utf-8')).hexdigest()

# =========================
# Main (somente Google Sheets em append)
# =========================
def main():
    print("🚀 Iniciando scraper de notícias...")
    driver = setup_driver()

    resultados_por_termo = []
    resumo_coletas = []

    try:
        for termo in tqdm(SEARCH_TERMS, desc="🔎 Buscando termos"):
            noticias = scrape_news_for_term(driver, termo)
            df_noticias = pd.DataFrame(noticias)
            df_noticias_24h = filter_recent_24h(df_noticias)
            resultados_por_termo.append(df_noticias_24h)
            resumo_coletas.append({'Termo de Busca': termo, 'Notícias Coletadas (24h)': len(df_noticias_24h)})

        df_geral = pd.concat(resultados_por_termo, ignore_index=True) if resultados_por_termo else pd.DataFrame()

        # Dedup local por URL Final (ou Link)
        if not df_geral.empty:
            if 'URL Final' in df_geral.columns:
                df_geral = df_geral.drop_duplicates(subset=['URL Final']).copy()
            else:
                df_geral = df_geral.drop_duplicates(subset=['Link']).copy()

        coletado_em_str = fmt_brasilia(agora_brasilia())
        if not df_geral.empty:
            df_geral['Coletado em'] = coletado_em_str

        # Dedup remota via fingerprints
        new_fps = []
        df_para_append = pd.DataFrame()
        try:
            gc = _gspread_client_from_env()
            _ = gc.open_by_key(SPREADSHEET_KEY)  # sanity check / permissão
        except Exception as e:
            print(f"❌ Erro com credenciais/planilha: {e}")
            sys.exit(1)

        if not df_geral.empty:
            df_geral['__fp'] = df_geral.apply(make_fingerprint, axis=1)
            try:
                existing = load_fingerprints(gc)
                mask_novos = ~df_geral['__fp'].isin(existing)
                df_para_append = df_geral[mask_novos].copy()
                new_fps = df_para_append['__fp'].tolist()
            except Exception as e:
                print(f"⚠️ Não foi possível usar fingerprints remotos (seguindo sem dedup remota): {e}")
                df_para_append = df_geral.copy()

            if '__fp' in df_para_append.columns:
                df_para_append.drop(columns=['__fp'], inplace=True)

        # APPEND no Google Sheets
        sheets_ok = False
        try:
            if not df_para_append.empty:
                append_news_rows(gc, df_para_append)
            else:
                print("ℹ️ Nenhuma notícia nova após deduplicação por fingerprints.")
            if new_fps:
                append_fingerprints(gc, new_fps)
            sheets_ok = True
        except Exception as e:
            print(f"❌ Erro ao fazer append no Google Sheets: {e}")
            sheets_ok = False

        total_noticias = len(df_para_append) if not df_para_append.empty else 0
        print(f"\n📊 RESUMO DA EXECUÇÃO:")
        print(f"📰 Total de notícias novas (24h, sem repetição): {total_noticias}")
        print(f"🕒 Coletado em (BRT): {coletado_em_str}")
        print(f"📤 Google Sheets (append): {'✅' if sheets_ok else '❌'}")
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
