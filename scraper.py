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
SPREADSHEET_KEY = (
    os.getenv("SPREADSHEET_KEY")
    or os.getenv("PLANILHA")  # aceita secret chamado "planilha"
)

TAB_NOTICIAS = "google notícias"   # única aba usada pelo projeto
TAB_CONFIG  = "Config"             # opcional (colunas: Termo, Ativo)

HTTP_TIMEOUT = 12
USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119 Safari/537.36"

# colunas exportadas para a planilha (sem Fingerprint nem histórico)
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
    if not SPREADSHEET_KEY:
        raise RuntimeError("Defina o secret SPREADSHEET_KEY ou PLANILHA com o ID da planilha.")
    gc = _gspread_client_from_env()
    return gc.open_by_key(SPREADSHEET_KEY)

# =========================
# SELENIUM
# =========================
def setup_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920x1080")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-plugins")
    # não desabilitar JS: GNews precisa de JS

    chrome_path = shutil.which("google-chrome") or shutil.which("chrome") or shutil.which("chromium")
    if chrome_path:
        options.binary_location = chrome_path

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
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
    """domínio + path (sem query/fragment) para dedup estável."""
    try:
        p = urlparse(url)
        base = urlunparse((p.scheme, p.netloc, p.path.rstrip("/"), "", "", ""))
        return base.lower()
    except Exception:
        return url.lower()

def resolve_final_url_requests(url: str) -> str:
    """Fallback via requests (segue redirects + tenta canonical/meta refresh)."""
    try:
        r = requests.get(url, timeout=HTTP_TIMEOUT, allow_redirects=True, headers={"User-Agent": USER_AGENT})
        final = r.url
        if "news.google.com" in urlparse(final).netloc:
            soup = BeautifulSoup(r.text, "html.parser")
            can = soup.find("link", rel=lambda v: v and "canonical" in v.lower())
            if can and can.get("href"):
                final = can["href"]
            else:
                meta = soup.find("meta", attrs={"http-equiv": lambda v: v and v.lower() == "refresh"})
                if meta and "url=" in meta.get("content","").lower():
                    final = meta["content"].split("url=",1)[-1].strip()
        return strip_tracking_params(final)
    except Exception:
        return strip_tracking_params(url)

def resolve_final_url_selenium(driver, url: str) -> str:
    """Abre o link do Google News numa nova aba e captura a URL destino real."""
    try:
        orig = driver.current_window_handle
        driver.switch_to.new_window("tab")
        driver.get(url)
        time.sleep(2.5)  # espera redirects/client-side
        final = driver.current_url
        for _ in range(2):  # pequenas esperas extras
            time.sleep(1.5)
            cur = driver.current_url
            if cur != final:
                final = cur
        driver.close()
        driver.switch_to.window(orig)
        if "news.google.com" in urlparse(final).netloc:
            final = resolve_final_url_requests(final)
        return strip_tracking_params(final)
    except Exception:
        try:
            for h in driver.window_handles:
                driver.switch_to.window(h)
                break
        except Exception:
            pass
        return resolve_final_url_requests(url)

def fetch_publisher_title(final_url: str) -> str:
    """Tenta pegar o título do site de origem (og:title / twitter:title / <title>)."""
    try:
        r = requests.get(final_url, timeout=HTTP_TIMEOUT, headers={"User-Agent": USER_AGENT})
        r.raise_for_status()
        s = BeautifulSoup(r.text, "html.parser")
        for sel in ['meta[property="og:title"]', 'meta[name="twitter:title"]']:
            m = s.select_one(sel)
            if m and m.get("content"):
                return m["content"].strip()
        if s.title and s.title.string:
            return s.title.string.strip()
    except Exception:
        pass
    return ""

# =========================
# SCRAPER
# =========================
def scrape_news_for_term(driver, termo, coletado_em_str: str):
    q = termo.replace(" ", "+")
    link = f"https://news.google.com/search?q={q}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"
    driver.get(link)
    time.sleep(3)

    # scroll até carregar tudo
    last_h = driver.execute_script("return document.body.scrollHeight")
    for _ in range(12):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h == last_h:
            break
        last_h = new_h

    soup = BeautifulSoup(driver.page_source, "html.parser")
    # preferir artigos com h3/h4 -> a (evita tópicos/edições)
    items = soup.select("article:has(h3 a), article:has(h4 a)")
    base = "https://news.google.com"

    rows = []
    for it in items:
        try:
            h = it.find(["h3", "h4"])
            a = h.find("a", href=True) if h else None
            title = a.get_text(strip=True) if a else (h.get_text(strip=True) if h else None)
            href = a["href"] if a else None
            if not href:
                continue

            # normaliza link relativo do GNews
            if href.startswith("./"):
                href = base + href[1:]
            elif href.startswith("/"):
                href = base + href

            # descarte páginas de cluster do próprio GNews
            parsed = urlparse(href)
            if "news.google.com" in parsed.netloc and any(
                seg in parsed.path for seg in ("/topics", "/publications", "/headlines", "/stories")
            ):
                continue

            # publisher/fonte
            pub = it.find("div", class_="vr1PYe") or it.find("div", class_="wsLqz")
            fonte = pub.get_text(strip=True) if pub else ""

            # data (string dd/mm/aaaa)
            t = it.find("time")
            dt_pub = ""
            if t and t.get("datetime"):
                try:
                    dt = datetime.strptime(t["datetime"], "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    dt_pub = dt.astimezone().strftime("%d/%m/%Y")
                except Exception:
                    dt_pub = ""

            # URL final real
            url_final = resolve_final_url_selenium(driver, href)
            dominio   = urlparse(url_final).netloc if url_final else ""
            fp        = url_fingerprint(url_final)

            # Corrigir título se for genérico do cluster
            needs_real_title = (
                not title
                or title.lower().startswith("notícias sobre")
                or (" • " in title and "notícias sobre" in title.lower())
                or "news.google.com" in urlparse(url_final).netloc
            )
            if needs_real_title:
                real = fetch_publisher_title(url_final)
                if real:
                    title = real
            if not title:
                continue  # evita linha feia

            rows.append({
                "URL Final": url_final,
                "URL GoogleNews": href,
                "Título": title,
                "Fonte": fonte,
                "Termo": termo,
                "Data de Publicação": dt_pub,
                "Dominio": dominio,
                "Coletado em": coletado_em_str,
                # campos só internos (não exportados)
                "_Fingerprint": fp
            })
        except Exception:
            continue

    return pd.DataFrame(rows)

# — filtro últimas 24h (robusto para tz-aware)
def filter_last_24h(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    d = pd.to_datetime(df["Data de Publicação"], format="%d/%m/%Y", errors="coerce")
    now = pd.Timestamp.utcnow()
    if now.tzinfo is None:
        now = now.tz_localize("UTC")
    try:
        d = d.dt.tz_localize("UTC")
    except (TypeError, AttributeError):
        d = d.dt.tz_convert("UTC")
    mask = (d.isna()) | (now.normalize() - d <= pd.Timedelta(days=1))
    return df[mask].copy()

# =========================
# SHEETS I/O (uma aba + append de novos)
# =========================
def read_terms_from_config(sh):
    try:
        ws = sh.worksheet(TAB_CONFIG)
        dfc = pd.DataFrame(ws.get_all_records())
        if dfc.empty or "Termo" not in dfc.columns:
            raise ValueError
        if "Ativo" in dfc.columns:
            dfc = dfc[dfc["Ativo"].astype(str).str.lower().isin(["true","1","sim","yes","y"])]
        termos = [t for t in dfc["Termo"].astype(str).str.strip().tolist() if t]
        if termos:
            return termos
    except Exception:
        pass
    return ['PNE', 'Plano Nacional de Educação', 'Saúde Mental']

def ensure_sheet(sh):
    try:
        return sh.worksheet(TAB_NOTICIAS)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=TAB_NOTICIAS, rows="2000", cols="30")

def normalize_missing_cols(df: pd.DataFrame, required_cols) -> pd.DataFrame:
    df = df.copy()
    for c in required_cols:
        if c not in df.columns:
            df[c] = ""
    return df

def upsert_append_only(sh, df_day: pd.DataFrame):
    """
    Acrescenta apenas itens cujo fingerprint (derivado da URL Final) ainda não exista na planilha.
    Preserva colunas manuais já existentes.
    """
    ws = ensure_sheet(sh)
    existing = pd.DataFrame(ws.get_all_records())
    existing = normalize_missing_cols(existing, EXPORT_COLS)

    # set de fingerprints já presentes (derivados da URL Final salva)
    if existing.empty:
        seen_fps = set()
    else:
        try:
            seen_fps = set(existing["URL Final"].map(url_fingerprint).astype(str).tolist())
        except Exception:
            seen_fps = set()

    # separa novos (com base no fingerprint interno)
    df_day = df_day.copy()
    df_day["_Fingerprint"] = df_day["_Fingerprint"].astype(str)
    novos = df_day[~df_day["_Fingerprint"].map(lambda x: x in seen_fps)].copy()

    # prepara saída: só colunas exportadas
    for c in EXPORT_COLS:
        if c not in novos.columns:
            novos[c] = ""
    novos = novos[EXPORT_COLS]

    if existing.empty or not set(EXPORT_COLS).issubset(existing.columns):
        out = novos.copy()
        # Ordena mais recentes no topo
        parsed = pd.to_datetime(out["Coletado em"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
        out = out.assign(_ord=parsed).sort_values("_ord", ascending=False, na_position="last").drop(columns=["_ord"])
        set_with_dataframe(ws, out)
        print(f"🧾 Inseridos {len(out)} itens (primeira carga).")
        return

    # preserva colunas manuais
    manual_cols = [c for c in existing.columns if c not in EXPORT_COLS]
    for c in manual_cols:
        if c not in novos.columns:
            novos[c] = ""

    ordered_cols = EXPORT_COLS + manual_cols + [c for c in existing.columns if c not in EXPORT_COLS + manual_cols]
    for c in ordered_cols:
        if c not in existing.columns:
            existing[c] = ""

    existing = existing[ordered_cols]
    updated = pd.concat([existing, novos[ordered_cols]], ignore_index=True)

    # ordena por coleta desc
    parsed = pd.to_datetime(updated["Coletado em"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    updated = updated.assign(_ord=parsed).sort_values("_ord", ascending=False, na_position="last").drop(columns=["_ord"])

    set_with_dataframe(ws, updated)
    print(f"✅ Append | novos: {len(novos)} | total: {len(updated)}")

# =========================
# MAIN
# =========================
def main():
    sh = open_sheet()
    termos = read_terms_from_config(sh)
    coletado_em_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    driver = setup_driver()
    try:
        dfs = []
        for termo in termos:
            df = scrape_news_for_term(driver, termo, coletado_em_str)
            df = filter_last_24h(df)
            dfs.append(df)

        df_all = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        if df_all.empty:
            print("⚠️ Nenhuma notícia encontrada nas últimas 24h.")
            return

        # dedup dentro da própria execução (mesmo artigo/termo/fonte)
        for c in ["_Fingerprint","Termo","Fonte","Título"]:
            if c not in df_all.columns:
                df_all[c] = ""
        df_all = df_all.drop_duplicates(subset=["_Fingerprint","Termo","Fonte","Título"])

        # garante colunas exportadas
        for c in EXPORT_COLS:
            if c not in df_all.columns:
                df_all[c] = ""

        upsert_append_only(sh, df_all)
    finally:
        driver.quit()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        sys.exit(1)
