import os
import re
import time
import hashlib
from datetime import datetime, timedelta, timezone
from urllib.parse import urlparse

import feedparser
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials


def _now_local():
    tz_offset = int(os.getenv("TZ_OFFSET_HOURS", "-3"))
    return datetime.now(timezone(timedelta(hours=tz_offset)))


def normalizar_url(url):
    if not url:
        return ""
    u = url.strip()
    u = re.sub(r"#.*$", "", u)
    return u.rstrip("/")


def extrair_dominio_e_path(link):
    p = urlparse(link)
    dominio = p.netloc.replace("www.", "")
    path = p.path.rstrip("/")
    return dominio, path


def google_news_rss_query(link):
    dominio, path = extrair_dominio_e_path(link)
    if path and path != "":
        return f"https://news.google.com/rss/search?q=site:{dominio}{path}"
    return f"https://news.google.com/rss/search?q=site:{dominio}"


def ler_sites(sheet):
    records = sheet.get_all_records()
    df = pd.DataFrame(records)
    if df.empty:
        raise ValueError("A aba de entrada está vazia.")
    df.columns = [c.strip().lower() for c in df.columns]
    if "link" not in df.columns or "estado" not in df.columns:
        raise ValueError("Colunas obrigatórias: Link, Estado")
    return df


def conectar_planilha(spreadsheet_id):
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(spreadsheet_id)


def garantir_aba(ss, nome):
    try:
        return ss.worksheet(nome)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=nome, rows=2000, cols=10)
        return ws


def garantir_header(ws):
    values = ws.get_all_values()
    if len(values) >= 2:
        return
    ws.insert_rows(
        [
            [],
            ["Data", "Estado", "Fonte", "Título", "URL"]
        ],
        row=1
    )


def ler_urls_existentes(ws):
    try:
        col = ws.col_values(5)
        return set(normalizar_url(u) for u in col if u)
    except Exception:
        return set()


def coletar_feed(feed_url, estado, lookback_dias):
    feed = feedparser.parse(feed_url)
    resultados = []

    limite = _now_local() - timedelta(days=lookback_dias)

    for e in feed.entries:
        link = normalizar_url(e.get("link"))
        if not link:
            continue

        dt_pub = None
        if hasattr(e, "published_parsed") and e.published_parsed:
            dt_pub = datetime(*e.published_parsed[:6], tzinfo=timezone.utc).astimezone(
                _now_local().tzinfo
            )

        if dt_pub and dt_pub < limite:
            continue

        titulo = e.get("title", "").strip()
        fonte = feed.feed.get("title", "")

        resultados.append({
            "data": dt_pub.strftime("%Y-%m-%d %H:%M") if dt_pub else "",
            "estado": estado,
            "fonte": fonte,
            "titulo": titulo,
            "url": link,
        })

    return resultados


def coletar_por_uf():
    spreadsheet_id = os.getenv("PLANILHA_SUBNACIONAL", "").strip()
    lookback_dias = int(os.getenv("LOOKBACK_DIAS", "1"))
    pausa = float(os.getenv("PAUSA_SEG", "1.2"))

    if not spreadsheet_id:
        raise ValueError("PLANILHA_SUBNACIONAL não definido")

    ss = conectar_planilha(spreadsheet_id)

    aba_entrada = ss.worksheet("deduplicado")
    df_sites = ler_sites(aba_entrada)

    for estado in sorted(df_sites["estado"].unique()):
        ws = garantir_aba(ss, estado.upper())
        garantir_header(ws)
        existentes = ler_urls_existentes(ws)

        novos = []

        subset = df_sites[df_sites["estado"] == estado]
        for _, row in subset.iterrows():
            link = row["link"]
            feed_url = google_news_rss_query(link)
            entradas = coletar_feed(feed_url, estado.upper(), lookback_dias)

            for r in entradas:
                if r["url"] in existentes:
                    continue
                existentes.add(r["url"])
                novos.append([
                    r["data"],
                    r["estado"],
                    r["fonte"],
                    r["titulo"],
                    r["url"],
                ])

            time.sleep(pausa)

        if novos:
            ws.insert_rows(novos, row=3)
            print(f"✓ {estado.upper()}: inseridas {len(novos)} linhas.")


if __name__ == "__main__":
    coletar_por_uf()
