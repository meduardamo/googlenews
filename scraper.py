# -*- coding: utf-8 -*-
import os
import json
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By  # mantido por compatibilidade
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from datetime import datetime
from tqdm import tqdm
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

# =======================
# CONFIGURAÇÕES SELENIUM
# =======================
options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920x1080")
options.add_argument("--blink-settings=imagesEnabled=false")

# garante binário do Chrome nos runners CI
chrome_path = shutil.which("google-chrome") or shutil.which("chrome") or shutil.which("chromium")
if chrome_path:
    options.binary_location = chrome_path

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# =======================
# PARÂMETROS DE BUSCA
# =======================
search_terms = ['PNE', 'Plano Nacional de Educação', 'saúde mental']
root = 'https://news.google.com'

resultados_por_termo = []
resumo_coletas = []

for termo in tqdm(search_terms, desc="🔎 Buscando termos"):
    print(f"\n🔍 Buscando notícias para: {termo}")

    query_text = termo.replace(' ', '+')
    link = f"https://news.google.com/search?q={query_text}&hl=pt-BR&gl=BR&ceid=BR%3Apt-419"

    driver.get(link)
    time.sleep(3)

    # Scroll até o fim
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            print("📜 Final do conteúdo alcançado.")
            break
        last_height = new_height

    # Pega o HTML carregado
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    news_items = soup.select('div.UW0SDc, article')

    noticias = []

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
                    datetime_obj = datetime.strptime(datetime_string, '%Y-%m-%dT%H:%M:%SZ')
                    data_publicacao = datetime_obj.strftime('%d/%m/%Y')
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
            print(f"Erro ao processar notícia: {e}")
            continue

    # Criar DataFrame do termo
    df_noticias = pd.DataFrame(noticias)

    # 🔥 Filtrar notícias das últimas 24h
    hoje = pd.Timestamp.now()
    if not df_noticias.empty:
        df_noticias['Data Convertida'] = pd.to_datetime(df_noticias['Data de Publicação'], format='%d/%m/%Y', errors='coerce')
        df_noticias['Dias de Diferença'] = (hoje - df_noticias['Data Convertida']).dt.days
        df_noticias_24h = df_noticias[df_noticias['Dias de Diferença'] <= 1].copy()
    else:
        df_noticias_24h = pd.DataFrame()

    print(f"✅ {len(df_noticias_24h)} notícias coletadas para o termo '{termo}'.")

    resultados_por_termo.append(df_noticias_24h)
    resumo_coletas.append({'Termo de Busca': termo, 'Notícias Coletadas': len(df_noticias_24h)})

# Fecha o navegador
driver.quit()

# Juntar todos os resultados
df_geral = pd.concat(resultados_por_termo, ignore_index=True) if resultados_por_termo else pd.DataFrame()

# Remover duplicadas (baseado no Link)
if not df_geral.empty and 'Link' in df_geral.columns:
    df_geral = df_geral.drop_duplicates(subset='Link').copy()

# Remover colunas técnicas antes de salvar
cols_drop = [c for c in ['Data Convertida', 'Dias de Diferença'] if c in df_geral.columns]
df_geral_final = df_geral.drop(columns=cols_drop) if not df_geral.empty else pd.DataFrame()

# Salvar localmente no Excel
with pd.ExcelWriter('noticias_PNE_Planos_Educacao.xlsx') as writer:
    (df_geral_final if not df_geral_final.empty else pd.DataFrame(columns=['Título','Fonte','Data de Publicação','Link','Termo de Busca']))\
        .to_excel(writer, sheet_name='Noticias', index=False)
    df_resumo = pd.DataFrame(resumo_coletas)
    df_resumo.to_excel(writer, sheet_name='Resumo', index=False)

print("\n✅ Notícias salvas no Excel 'noticias_PNE_Planos_Educacao.xlsx'.")

# =======================
# GOOGLE SHEETS (via SECRET)
# =======================
# Requer secret GCP_SERVICE_ACCOUNT_JSON com o JSON completo da conta de serviço
service_account_json = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")
if not service_account_json:
    raise RuntimeError("Secret GCP_SERVICE_ACCOUNT_JSON não encontrado nas variáveis de ambiente.")

service_account_info = json.loads(service_account_json)
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scopes=scope)
gc = gspread.authorize(credentials)

# Abrir a planilha por KEY (ajuste se necessário)
SPREADSHEET_KEY = '1G81BndSPpnViMDxRKQCth8PwK0xmAwH-w-T7FjgnwcY'
worksheet_name = 'PNE'

spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except Exception:
    worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="100", cols="20")

set_with_dataframe(
    worksheet,
    df_geral_final if not df_geral_final.empty else pd.DataFrame(columns=['Título','Fonte','Data de Publicação','Link','Termo de Busca'])
)

print("\n✅ Dados enviados para o Google Sheets na aba 'PNE'.")