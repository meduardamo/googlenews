# -*- coding: utf-8 -*-
import os, json, time, shutil, sys, requests
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
SPREADSHEET_KEY = os.getenv("SPREADSHEET_KEY") or os.getenv("PLANILHA")
TAB_NOTICIAS = "google notícias"       # única aba usada
TAB_CONFIG  = "Config"                 # opcional (colunas: Termo, Ativo)
WINDOW_DAYS = 1                        # últimas 24h (mude para 7, 30, etc. se quiser)

HTTP_TIMEOUT = 12
USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119 Safari/537.36"

EXPORT_COLS = [
    "URL Final","URL GoogleNews","Título","Fonte",
    "Termo","Data de Publicação","Dominio","Coletado em"
]

# =========================
# AUTH GOOGLE SHEETS
# =========================
def _gspread_client_from_env():
    raw = os.environ.get("GCP_SERVICE_ACCOUNT_JSON") or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if not raw:
        raise RuntimeError("Secret GCP_SERVICE_ACCOUNT_JSON/GOOGLE_APPLICATION_CREDENTIALS_JSON não encontrado.")
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
    if not SPREADSHEET_KEY:
        raise RuntimeError("Defina SPREADSHEET_KEY ou PLANILHA com o ID da planilha.")
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
    # ⚠️ não desabilitar JS (GNews precisa)

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
# URL helpers
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

def url_fingerprint(url: str) -> str:
    # domínio + path, sem query/fragment
    try:
        p = urlparse(url)
        base = urlunparse((p.scheme, p.netloc, p.path.rstrip("/"), "", "", ""))
        return base.lower()
    except Exception:
        return url.lower()

def resolve_final_url_requests(url: str) -> str:
    # fallback via requests: redirects + canonical/meta refresh
    try:
        r = requests.get(url, timeout=HTTP_TIMEOUT, allow_redirects=True, headers={"User-Agent": USER_AGENT})
        final = r.url
        if "news.google.com" in urlparse(final).netloc:
            s = BeautifulSoup(r.text, "html.parser")
            can = s.find("link", rel=lambda v: v and "canonical" in v.lower())
            if can and can.get("href"):
                final = can["href"]
            else:
                meta = s.find("meta", attrs={"http-equiv": lambda v: v and v.lower() == "refresh"})
                if meta and "url=" in (meta.get("content","").lower()):
                    final = meta["content"].split("url=",1)[-1].strip()
        return strip_tracking_params(final)
    except Exception:
        return strip_tracking_params(url)

def resolve_final_url_selenium(driver, url: str) -> str:
    # abre em nova aba e captura o destino real
    try:
        orig = driver.current_window_handle
        driver.switch_to.new_window("tab")
        driver.get(url)
        time.sleep(2.5)
        final = driver.current_url
        for _ in range(2):
            time.sleep(1.5)
            cur = driver.current_url
            if cur != final: final = cur
        driver.close()
        driver.switch_to.window(orig)
        if "news.google.com" in urlparse(final).netloc:
            final = resolve_final_url_requests(final)
        return strip_tracking_params(final)
    except Exception:
        try:
            for h in driver.window_handles:
                driver.switch_to.window(h); break
        except Exception:
            pass
        return resolve_final_url_requests(url)

def fetch_publisher_title(final_url: str) -> str:
    try:
        r = requests.get(final_url, timeout=HTTP_TIMEOUT, headers={"User-Agent": USER_AGENT})
        r.raise_for_status()
        s = BeautifulSoup(r.text, "html.parser")
        for sel in ['meta[property="og:title"]','meta[name="twitter:title"]']:
            m = s.select_one(sel)
            if m and m.get("content"):
                return m["content"].strip()
        if s.title and s.title.string:
            return s.title.string.strip()
    except Exception:
        pass
    return ""

# =========================
# BUSCA (24h) + SCRAPE
# =========================
def scrape_news_for_term(driver, termo, coletado_em_str: str):
    # força período no próprio Google News (24h)
    termo_periodo = f"{termo} when:{WINDOW_DAYS}d"
    q = termo_periodo.replace(' ', '+')
    link = f"https://news.google.com/search?q={q}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"

    print(f"\n🔍 Termo: {termo} | Período: {WINDOW_DAYS}d")
    driver.get(link)
    time.sleep(3)

    # scroll para carregar resultados
    last_h = driver.execute_script("return document.body.scrollHeight")
    for _ in range(12):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h == last_h: break
        last_h = new_h

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # pega cartões de artigo; mantemos seu seletor original como primeira opção
    items = soup.select('div.UW0SDc, article')

    base = 'https://news.google.com'
    rows = []

    for it in items:
        try:
            # === TÍTULO (como no seu código original) ===
            title_el = it.find('a', class_='JtKRv') or it.find('h3') or it.find('h4')
            title = title_el.get_text(strip=True) if title_el else None

            # link (normaliza relativo do GNews)
            a = it.find("a", href=True)
            href = a['href'] if a else None
            if not href: continue
            if href.startswith('./'): href = base + href[1:]
            elif href.startswith('/'): href = base + href

            # filtra páginas de cluster/tópico do GNews
            p = urlparse(href)
            if "news.google.com" in p.netloc and any(seg in p.path for seg in ("/topics","/publications","/headlines","/stories")):
                continue

            # fonte
            pub = it.find('div', class_='vr1PYe') or it.find('div', class_='wsLqz')
            fonte = pub.get_text(strip=True) if pub else ""

            # data
            t = it.find('time', class_='hvbAAd') or it.find('time')
            dt_pub = ""
            if t and t.get('datetime'):
                try:
                    dt = datetime.strptime(t['datetime'], '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc)
                    dt_pub = dt.astimezone().strftime('%d/%m/%Y')
                except Exception:
                    dt_pub = ""

            # URL final real (abre e segue redirect)
            url_final = resolve_final_url_selenium(driver, href)
            dominio   = urlparse(url_final).netloc if url_final else ""
            fp        = url_fingerprint(url_final)

            # corrige títulos “Notícias sobre …” com título do publisher
            if (not title) or title.lower().startswith("notícias sobre"):
                real = fetch_publisher_title(url_final)
                if real: title = real
            if not title:  # se ainda não tiver, pula
                continue

            rows.append({
                "URL Final": url_final,
                "URL GoogleNews": href,
                "Título": title,
                "Fonte": fonte,
                "Termo": termo,
                "Data de Publicação": dt_pub,
                "Dominio": dominio,
                "Coletado em": coletado_em_str,
                "_Fingerprint": fp  # só interno
            })
        except Exception:
            continue

    return pd.DataFrame(rows)

# =========================
# FILTER (rede de segurança 24h)
# =========================
def filter_last_window(df: pd.DataFrame, days: int) -> pd.DataFrame:
    if df.empty: return df
    d = pd.to_datetime(df["Data de Publicação"], format="%d/%m/%Y", errors="coerce")
    now = pd.Timestamp.utcnow()
    if now.tzinfo is None: now = now.tz_localize("UTC")
    try: d = d.dt.tz_localize("UTC")
    except (TypeError, AttributeError): d = d.dt.tz_convert("UTC")
    return df[(d.isna()) | (now.normalize() - d <= pd.Timedelta(days=days))].copy()

# =========================
# SHEETS I/O (append-only)
# =========================
def read_terms_from_config(sh):
    try:
        ws = sh.worksheet(TAB_CONFIG)
        dfc = pd.DataFrame(ws.get_all_records())
        if dfc.empty or "Termo" not in dfc.columns: raise ValueError
        if "Ativo" in dfc.columns:
            dfc = dfc[dfc["Ativo"].astype(str).str.lower().isin(["true","1","sim","yes","y"])]
        termos = [t for t in dfc["Termo"].astype(str).str.strip().tolist() if t]
        if termos: return termos
    except Exception:
        pass
    return ['PNE', 'Plano Nacional de Educação', 'Saúde Mental']

def ensure_sheet(sh):
    try: return sh.worksheet(TAB_NOTICIAS)
    except gspread.WorksheetNotFound: return sh.add_worksheet(title=TAB_NOTICIAS, rows="2000", cols="30")

def normalize_missing_cols(df: pd.DataFrame, cols) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c not in df.columns: df[c] = ""
    return df

def upsert_append_only(sh, df_day: pd.DataFrame):
    ws = ensure_sheet(sh)
    existing = pd.DataFrame(ws.get_all_records())
    existing = normalize_missing_cols(existing, EXPORT_COLS)

    # fingerprints já presentes (derivados da URL Final salva)
    if existing.empty:
        seen = set()
    else:
        try: seen = set(existing["URL Final"].map(url_fingerprint).astype(str).tolist())
        except Exception: seen = set()

    # novos = os que não têm fingerprint existente
    df_day = df_day.copy()
    novos = df_day[~df_day["_Fingerprint"].map(lambda x: x in seen)].copy()

    for c in EXPORT_COLS:
        if c not in novos.columns: novos[c] = ""
    novos = novos[EXPORT_COLS]

    if existing.empty or not set(EXPORT_COLS).issubset(existing.columns):
        out = novos.copy()
        parsed = pd.to_datetime(out["Coletado em"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
        out = out.assign(_ord=parsed).sort_values("_ord", ascending=False, na_position="last").drop(columns=["_ord"])
        set_with_dataframe(ws, out)
        print(f"🧾 Inseridos {len(out)} itens (primeira carga).")
        return

    # preserva colunas manuais já existentes
    manual_cols = [c for c in existing.columns if c not in EXPORT_COLS]
    for c in manual_cols:
        if c not in novos.columns: novos[c] = ""

    ordered = EXPORT_COLS + manual_cols + [c for c in existing.columns if c not in EXPORT_COLS + manual_cols]
    for c in ordered:
        if c not in existing.columns: existing[c] = ""
    existing = existing[ordered]

    updated = pd.concat([existing, novos[ordered]], ignore_index=True)
    parsed = pd.to_datetime(updated["Coletado em"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    updated = updated.assign(_ord=parsed).sort_values("_ord", ascending=False, na_position="last").drop(columns=["_ord"])

    set_with_dataframe(ws, updated)
    print(f"✅ Append | novos: {len(novos)} | total: {len(updated)}")

# =========================
# MAIN
# =========================
def main():
    print("🚀 Iniciando scraper Google News (24h)")
    sh = open_sheet()
    termos = read_terms_from_config(sh)
    coletado_em_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    driver = setup_driver()
    try:
        dfs = []
        for termo in termos:
            df = scrape_news_for_term(driver, termo, coletado_em_str)
            df = filter_last_window(df, WINDOW_DAYS)   # rede de segurança
            dfs.append(df)

        df_all = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        if df_all.empty:
            print("⚠️ Nenhuma notícia encontrada no período (verifique termos/horário).")
            return

        # dedup dentro da execução
        for c in ["_Fingerprint","Termo","Fonte","Título"]:
            if c not in df_all.columns: df_all[c] = ""
        df_all = df_all.drop_duplicates(subset=["_Fingerprint","Termo","Fonte","Título"])

        # garante colunas exportadas
        for c in EXPORT_COLS:
            if c not in df_all.columns: df_all[c] = ""

        upsert_append_only(sh, df_all)
        print(f"📌 Coletadas e anexadas: {len(df_all)}")
    finally:
        driver.quit()
        print("🏁 Finalizado")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        sys.exit(1)
