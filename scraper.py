# -*- coding: utf-8 -*-
import os
import re
import json
import time
import shutil
import sys
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timezone
from urllib.parse import urlparse, urlunparse, parse_qsl, urlencode

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials

# =========================
# CONFIG
# =========================
SPREADSHEET_KEY = os.getenv("SPREADSHEET_KEY", "1G81BndSPpnViMDxRKQCth8PwK0xmAwH-w-T7FjgnwcY")
TAB_NOTICIAS = "google notícias"  # única aba
TAB_CONFIG = "Config"             # opcional (coluna Termo, e Ativo TRUE/FALSE)

HTTP_TIMEOUT = 12
USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119 Safari/537.36"

CORE_COLS = [
    "Fingerprint","URL Final","URL GoogleNews","Título","Fonte","Termo",
    "Data de Publicação","Dominio","Primeiro Visto","Último Visto","Vezes Visto"
]

# =========================
# AUTH GOOGLE SHEETS
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

def open_sheet():
    gc = _gspread_client_from_env()
    return gc.open_by_key(SPREADSHEET_KEY)

# =========================
# SELENIUM
# =========================
def setup_driver():
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
    # ⚠️ não desabilitar JS (Google News precisa de JS)

    if os.getenv('GITHUB_ACTIONS'):
        options.add_argument("--disable-background-timer-throttling")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--disable-backgrounding-occluded-windows")

    chrome_path = shutil.which("google-chrome") or shutil.which("chrome") or shutil.which("chromium")
    if chrome_path:
        options.binary_location = chrome_path
        print(f"✅ Chrome: {chrome_path}")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    print("✅ WebDriver OK")
    return driver

# =========================
# URL helpers (limpar UTM / seguir redirect)
# =========================
UTM_PARAMS = {"utm_source","utm_medium","utm_campaign","utm_term","utm_content",
              "utm_id","utm_name","gclid","fbclid","mcid","ocid","igshid"}

def strip_tracking_params(url: str) -> str:
    try:
        p = urlparse(url)
        q = [(k, v) for k, v in parse_qsl(p.query, keep_blank_values=True) if k.lower() not in UTM_PARAMS]
        new_query = urlencode(q)
        clean = p._replace(query=new_query, fragment="")
        return urlunparse(clean)
    except Exception:
        return url

def resolve_final_url(url: str) -> str:
    """
    Segue o redirect do Google News (/articles/...) até o publisher.
    Retorna a URL final limpa (sem UTMs).
    """
    try:
        r = requests.get(url, timeout=HTTP_TIMEOUT, allow_redirects=True, headers={"User-Agent": USER_AGENT})
        final = r.url
        return strip_tracking_params(final)
    except Exception:
        return strip_tracking_params(url)

def url_fingerprint(url: str) -> str:
    """domínio + path, sem query/fragment (para dedup estável)"""
    try:
        p = urlparse(url)
        base = urlunparse((p.scheme, p.netloc, p.path.rstrip("/"), "", "", ""))
        return base.lower()
    except Exception:
        return url.lower()

# =========================
# SCRAPER
# =========================
def scrape_news_for_term(driver, termo):
    print(f"\n🔍 Termo: {termo}")
    q = termo.replace(" ", "+")
    link = f"https://news.google.com/search?q={q}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"
    driver.get(link)
    time.sleep(3)

    # scroll para carregar itens
    last_h = driver.execute_script("return document.body.scrollHeight")
    for _ in range(12):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h == last_h:
            break
        last_h = new_h

    soup = BeautifulSoup(driver.page_source, "html.parser")
    items = soup.select("div.UW0SDc, article")
    base = "https://news.google.com"

    rows = []
    for it in items:
        try:
            # título
            h = it.find(["h3", "h4"])
            title = h.get_text(strip=True) if h else None
            # link relativo
            a = it.find("a", href=True)
            href = a["href"] if a else None
            if href and href.startswith("./"):
                href = base + href[1:]
            elif href and href.startswith("/"):
                href = base + href
            # publisher
            pub = it.find("div", class_="vr1PYe") or it.find("div", class_="wsLqz")
            fonte = pub.get_text(strip=True) if pub else ""
            # data
            t = it.find("time")
            dt_pub = ""
            if t and t.get("datetime"):
                try:
                    dt = datetime.strptime(t["datetime"], "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    dt_pub = dt.astimezone().strftime("%d/%m/%Y")
                except Exception:
                    dt_pub = ""

            if not (title and href):
                continue

            url_final = resolve_final_url(href)
            fp = url_fingerprint(url_final)

            rows.append({
                "Título": title,
                "Fonte": fonte,
                "Data de Publicação": dt_pub,
                "URL GoogleNews": href,
                "URL Final": url_final,
                "Fingerprint": fp,
                "Termo": termo
            })
        except Exception:
            continue

    return pd.DataFrame(rows)

def filter_last_24h(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: 
        return df
    d = pd.to_datetime(df["Data de Publicação"], format="%d/%m/%Y", errors="coerce")
    now = pd.Timestamp.utcnow().tz_localize("UTC")
    mask = (d.isna()) | (now.normalize() - d.dt.tz_localize("UTC", nonexistent="shift_forward", ambiguous="NaT") <= pd.Timedelta(days=1))
    return df[mask].copy()

# =========================
# SHEETS: leitura de termos e UPSERT na mesma aba
# =========================
def read_terms_from_config(sh):
    try:
        ws = sh.worksheet(TAB_CONFIG)
        dfc = pd.DataFrame(ws.get_all_records())
        if dfc.empty or "Termo" not in dfc.columns:
            raise ValueError("Aba Config vazia ou sem coluna 'Termo'.")
        if "Ativo" in dfc.columns:
            dfc = dfc[dfc["Ativo"].astype(str).str.lower().isin(["true","1","sim","yes","y"])]
        termos = [t for t in dfc["Termo"].astype(str).str.strip().tolist() if t]
        if termos:
            print(f"🧩 Termos (Config): {termos}")
            return termos
    except Exception:
        pass
    print("➡️ Usando termos padrão: ['PNE','Plano Nacional de Educação','Saúde Mental']")
    return ['PNE', 'Plano Nacional de Educação', 'Saúde Mental']

def ensure_sheet(sh):
    try:
        return sh.worksheet(TAB_NOTICIAS)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=TAB_NOTICIAS, rows="2000", cols="30")

def prepare_new_rows(df_day: pd.DataFrame, coletado_em_str: str) -> pd.DataFrame:
    if df_day.empty:
        return df_day
    df = df_day.copy()
    # Dominio
    df["Dominio"] = df["URL Final"].apply(lambda u: urlparse(str(u)).netloc if pd.notna(u) else "")
    # métricas iniciais
    df["Primeiro Visto"] = coletado_em_str
    df["Último Visto"] = coletado_em_str
    df["Vezes Visto"] = 1
    for c in CORE_COLS:
        if c not in df.columns:
            df[c] = "" if c != "Vezes Visto" else 1
    return df[CORE_COLS + [c for c in df.columns if c not in CORE_COLS]]

def upsert_google_noticias(sh, df_day: pd.DataFrame, coletado_em_str: str):
    ws = ensure_sheet(sh)
    existing = pd.DataFrame(ws.get_all_records())
    new_rows = prepare_new_rows(df_day, coletado_em_str)

    if existing.empty:
        set_with_dataframe(ws, new_rows.sort_values("Último Visto", ascending=False))
        print(f"🧾 Inseridos {len(new_rows)} itens (primeira carga).")
        return

    # colunas “manuais” preservadas pela equipe (tudo que já existe e não é CORE)
    manual_cols = [c for c in existing.columns if c not in CORE_COLS]

    existing["Fingerprint"] = existing["Fingerprint"].astype(str)
    new_rows["Fingerprint"] = new_rows["Fingerprint"].astype(str)

    seen = set(existing["Fingerprint"])
    novos = new_rows[~new_rows["Fingerprint"].isin(seen)].copy()
    repetidos = new_rows[new_rows["Fingerprint"].isin(seen)].copy()

    # atualiza métricas dos repetidos
    if not repetidos.empty:
        rep_map = repetidos.set_index("Fingerprint")
        idx = existing["Fingerprint"].isin(rep_map.index)
        existing.loc[idx, "Último Visto"] = coletado_em_str
        existing.loc[idx, "Vezes Visto"] = pd.to_numeric(existing.loc[idx, "Vezes Visto"], errors="coerce").fillna(0).astype(int) + 1
        # opcional: atualizar campos vazios com info mais nova
        for col in ["URL Final","URL GoogleNews","Título","Fonte","Termo","Data de Publicação","Dominio"]:
            existing.loc[idx, col] = existing.loc[idx, col].where(existing.loc[idx, col].astype(str).str.len() > 0,
                                                                  rep_map.loc[existing.loc[idx, "Fingerprint"], col].values)

    # completar colunas manuais nos novos (vazias)
    for c in manual_cols:
        if c not in novos.columns:
            novos[c] = ""

    ordered_cols = CORE_COLS + manual_cols + [c for c in existing.columns if c not in CORE_COLS + manual_cols]
    for c in ordered_cols:
        if c not in existing.columns:
            existing[c] = ""
    existing = existing[ordered_cols]

    updated = pd.concat([existing, novos[ordered_cols]], ignore_index=True)

    # ordena por Último Visto desc
    if "Último Visto" in updated.columns:
        parsed = pd.to_datetime(updated["Último Visto"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
        updated = updated.assign(_ord=parsed).sort_values("_ord", ascending=False).drop(columns=["_ord"])

    set_with_dataframe(ws, updated)
    print(f"✅ Upsert | novos: {len(novos)} | atualizados: {len(repetidos)} | total: {len(updated)}")

# =========================
# MAIN
# =========================
def main():
    print("🚀 Iniciando scraper Google News")
    sh = open_sheet()
    termos = read_terms_from_config(sh)

    coletado_em_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    driver = setup_driver()
    try:
        dfs = []
        for termo in termos:
            df = scrape_news_for_term(driver, termo)
            df = filter_last_24h(df)
            dfs.append(df)

        df_all = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        if df_all.empty:
            print("⚠️ Nenhuma notícia encontrada (últimas 24h). Ainda assim mantendo histórico existente.")
            return

        # dedup por Fingerprint (mesmo artigo, múltiplos cards/termos)
        df_all["URL Final"] = df_all["URL Final"].fillna(df_all["URL GoogleNews"])
        df_all["Fingerprint"] = df_all["Fingerprint"].fillna(df_all["URL Final"]).astype(str)
        df_all = df_all[df_all["Fingerprint"].str.len() > 0]
        df_all = df_all.drop_duplicates(subset=["Fingerprint","Termo","Fonte","Título"])

        upsert_google_noticias(sh, df_all, coletado_em_str)
        print(f"📌 Coletados (24h, únicos por artigo/termo/fonte): {len(df_all)}")
    finally:
        driver.quit()
        print("🏁 Finalizado")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        sys.exit(1)
