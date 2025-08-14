# -*- coding: utf-8 -*-
"""
Raspagem Google News → Google Sheets → E-mail (Brevo)

Variáveis de ambiente esperadas:
- PLANILHA            (obrigatória)  -> key da planilha (ex: 1AbC...xYz)
- ABA                 (opcional)     -> nome da worksheet (default: "PNE")
- EMAIL               (obrigatória)  -> e-mail do remetente (precisa existir no Brevo)
- DESTINATARIOS       (obrigatória)  -> lista separada por vírgula (ex: "a@b.com,c@d.com")
- BREVO_API_KEY       (obrigatória)  -> API Key do Brevo
Arquivos:
- credentials.json    (obrigatório)  -> criado no CI a partir do secret GOOGLE_APPLICATION_CREDENTIALS_JSON
"""

import os
import json
import time
import shutil
import re
from datetime import datetime
from typing import Dict, List

import pandas as pd
from bs4 import BeautifulSoup
from tqdm import tqdm

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

# Brevo
from brevo_python import ApiClient, Configuration
from brevo_python.api.transactional_emails_api import TransactionalEmailsApi
from brevo_python.models.send_smtp_email import SendSmtpEmail


# ================================
# Configurações / Parâmetros
# ================================
SEARCH_TERMS = ['PNE', 'Plano Nacional de Educação', 'saúde mental']
GOOGLE_NEWS_ROOT = 'https://news.google.com'
GOOGLE_NEWS_QUERY = "https://news.google.com/search?q={query}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"
CREDENTIALS_PATH = "credentials.json"


# ================================
# Selenium (Chrome headless)
# ================================
def build_driver() -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920x1080")
    opts.add_argument("--blink-settings=imagesEnabled=false")

    chrome_path = shutil.which("google-chrome") or shutil.which("chrome") or shutil.which("chromium")
    if chrome_path:
        opts.binary_location = chrome_path

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    return driver


# ================================
# Raspagem Google News
# ================================
def raspa_google_news(search_terms: List[str]) -> Dict[str, pd.DataFrame]:
    driver = build_driver()
    resultados: Dict[str, pd.DataFrame] = {}

    for termo in tqdm(search_terms, desc="🔎 Buscando termos"):
        print(f"\n🔍 Buscando notícias para: {termo}")
        query = termo.replace(" ", "+")
        url = GOOGLE_NEWS_QUERY.format(query=query)

        driver.get(url)
        time.sleep(3)

        # Scroll até o fim para carregar tudo
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                print("📜 Final do conteúdo alcançado.")
                break
            last_height = new_height

        soup = BeautifulSoup(driver.page_source, "html.parser")
        news_items = soup.select("div.UW0SDc, article")

        noticias = []
        for item in news_items:
            try:
                title = item.find('a', class_='JtKRv') or item.find('h3') or item.find('h4')
                link_item = item.find("a", href=True)
                publisher = item.find('div', class_='vr1PYe') or item.find('div', class_='wsLqz')
                time_tag = item.find('time', class_='hvbAAd') or item.find('time')

                # Trata data
                data_publicacao = "Data não encontrada"
                if time_tag and time_tag.get("datetime"):
                    try:
                        dt = datetime.strptime(time_tag["datetime"], "%Y-%m-%dT%H:%M:%SZ")
                        data_publicacao = dt.strftime("%d/%m/%Y")
                    except ValueError:
                        data_publicacao = "Data inválida"

                link = "Link não encontrado"
                if link_item and link_item.get("href"):
                    href = link_item["href"]
                    if href.startswith("./"):
                        link = GOOGLE_NEWS_ROOT + href[1:]
                    elif href.startswith("/"):
                        link = GOOGLE_NEWS_ROOT + href
                    else:
                        link = href

                noticias.append({
                    "Título": (title.text.strip() if title else "Título não encontrado"),
                    "Fonte": (publisher.text.strip() if publisher else "Fonte não encontrada"),
                    "Data de Publicação": data_publicacao,
                    "Link": link,
                    "Termo de Busca": termo
                })
            except Exception as e:
                print(f"Erro ao processar notícia: {e}")

        df = pd.DataFrame(noticias)
        # Filtro últimas 24h
        hoje = pd.Timestamp.now()
        if not df.empty:
            df["Data Convertida"] = pd.to_datetime(df["Data de Publicação"], format="%d/%m/%Y", errors="coerce")
            df["Dias de Diferença"] = (hoje - df["Data Convertida"]).dt.days
            df = df[df["Dias de Diferença"] <= 1].copy()
            df.drop(columns=["Data Convertida", "Dias de Diferença"], inplace=True, errors="ignore")

        print(f"✅ {len(df)} notícias coletadas para '{termo}'.")
        resultados[termo] = df

    driver.quit()
    return resultados


# ================================
# Salvar Excel + Google Sheets
# ================================
def salva_resultados(resultados: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames = [df for df in resultados.values() if not df.empty]
    df_geral = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=["Título", "Fonte", "Data de Publicação", "Link", "Termo de Busca"]
    )
    if not df_geral.empty and "Link" in df_geral.columns:
        df_geral.drop_duplicates(subset="Link", inplace=True)

    # Excel local (para artifact do CI)
    writer_name = "noticias_PNE_Planos_Educacao.xlsx"
    with pd.ExcelWriter(writer_name) as w:
        df_geral.to_excel(w, sheet_name="Noticias", index=False)
        resumo = pd.DataFrame({
            "Termo de Busca": list(resultados.keys()),
            "Notícias Coletadas": [len(resultados[t]) for t in resultados],
        })
        resumo.to_excel(w, sheet_name="Resumo", index=False)
    print(f"\n✅ Excel salvo: {writer_name}")

    # Google Sheets
    spreadsheet_key = os.getenv("PLANILHA")
    if not spreadsheet_key:
        raise ValueError("❌ Variável de ambiente PLANILHA não definida (key da planilha)")

    aba = os.getenv("ABA", "PNE")

    scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    if not os.path.exists(CREDENTIALS_PATH):
        raise FileNotFoundError("❌ credentials.json não encontrado. No CI, ele é criado a partir do secret.")
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, scopes=scopes)
    gc = gspread.authorize(creds)

    ss = gc.open_by_key(spreadsheet_key)
    try:
        ws = ss.worksheet(aba)
    except Exception:
        ws = ss.add_worksheet(title=aba, rows="100", cols="20")

    set_with_dataframe(ws, df_geral)
    print(f"✅ Dados enviados ao Google Sheets ({spreadsheet_key}) na aba '{aba}'.")

    return df_geral


# ================================
# E-mail (Brevo)
# ================================
def envia_email_brevo(df_geral: pd.DataFrame, resultados: Dict[str, pd.DataFrame]) -> None:
    api_key = os.getenv("BREVO_API_KEY")
    remetente = os.getenv("EMAIL")
    destinatarios = os.getenv("DESTINATARIOS")

    if not api_key or not remetente or not destinatarios:
        print("⚠️ BREVO_API_KEY/EMAIL/DESTINATARIOS ausentes — pulando envio de e-mail.")
        return

    to_list = [e.strip() for e in destinatarios.split(",") if e.strip()]
    data = datetime.now().strftime("%d/%m/%Y")
    titulo = f"Google News ({data}) — PNE / Educação / Saúde Mental"

    planilha_key = os.getenv("PLANILHA")
    planilha_url = f"https://docs.google.com/spreadsheets/d/{planilha_key}/edit#gid=0" if planilha_key else "#"

    # Bloco de resumo por termo
    resumo_li = "".join(
        f"<li><b>{termo}</b>: {len(resultados[termo])} notícias</li>"
        for termo in resultados
    )

    # Algumas primeiras notícias (até 10)
    itens = []
    if not df_geral.empty:
        top = df_geral.head(10).to_dict("records")
        for r in top:
            itens.append(f"<li><a href='{r['Link']}' target='_blank'>{r['Título']}</a> — {r['Fonte']}</li>")
    itens_html = "".join(itens) if itens else "<li>(sem itens nas últimas 24h)</li>"

    html = f"""
    <html>
      <body>
        <h2>Coleta Google News — {data}</h2>
        <p>Resultados salvos na <a href="{planilha_url}" target="_blank">planilha</a>.</p>
        <h3>Resumo por termo</h3>
        <ul>{resumo_li}</ul>
        <h3>Alguns destaques</h3>
        <ul>{itens_html}</ul>
      </body>
    </html>
    """

    config = Configuration()
    config.api_key['api-key'] = api_key
    api = TransactionalEmailsApi(ApiClient(configuration=config))

    for dest in to_list:
        send_email = SendSmtpEmail(
            to=[{"email": dest}],
            sender={"email": remetente},
            subject=titulo,
            html_content=html
        )
        api.send_transac_email(send_email)
        print(f"✅ E-mail enviado para {dest}")


# ================================
# Execução principal
# ================================
if __name__ == "__main__":
    resultados = raspa_google_news(SEARCH_TERMS)
    df = salva_resultados(resultados)
    envia_email_brevo(df, resultados)
