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

WORKSHEET_FP = '_fingerprints'  # armazena fp e timestamp

HEADERS_DATA = ['Título','Fonte','Data de Publicação','Link','URL Final','Termo de Busca','Coletado em']
HEADERS_FP   = ['fp','created_at_brt']

# =========================
# Mapa de Clientes e Termos de Busca
# =========================
CLIENT_SEARCH_TERMS = {
    'IAS': ['Matemática', 'Alfabetização', 'Alfabetização Matemática', 'Recomposição de aprendizagem', 'Plano Nacional de Educação'],
    'ISG': ['Tempo Integral', 'Ensino em tempo integral', 'Ensino Profissional e Tecnológico', 'Fundeb', 'PROPAG', 'Educação em tempo integral', 'Escola em tempo integral', 'Plano Nacional de Educação', 'Programa escola em tempo integral', 'Programa Pé-de-meia', 'PNEERQ', 'INEP', 'FNDE', 'Conselho Nacional de Educação', 'PDDE', 'Programa de Fomento às Escolas de Ensino Médio em Tempo Integral', 'Celular nas escolas', 'Juros da Educação'],
    'IU': ['Gestão Educacional', 'Diretores escolares', 'Magistério', 'Professores ensino médio', 'Sindicatos de professores', 'Ensino Médio', 'Fundeb', 'Adaptações de Escolas', 'Educação Ambiental', 'Plano Nacional de Educação', 'PDDE', 'Programa Pé de Meia', 'INEP', 'FNDE', 'Conselho Nacional de Educação', 'VAAT', 'VAAR', 'Secretaria Estadual de Educação', 'Celular nas escolas', 'EAD', 'Juro da educação', 'Recomposição de Aprendizagem'],
    'Reúna': ['Matemática', 'Alfabetização', 'Alfabetização Matemática', 'Recomposição de aprendizagem', 'Plano Nacional de Educação', 'Emendas parlamentares educação'],
    'REMS': ['Esporte amador', 'Esporte para toda a vida', 'Esporte e desenvolvimento social', 'Financiamento do esporte', 'Lei de Incentivo ao Esporte', 'Plano Nacional de Esporte', 'Conselho Nacional de Esporte', 'Emendas parlamentares esporte'],
    'FMCSV': ['Criança', 'Infância', 'infanto-juvenil', 'educação básica', 'PNE', 'FNDE', 'Fundeb', 'VAAR', 'VAAT', 'educação infantil', 'maternidade', 'paternidade', 'alfabetização', 'creche', 'pré-escola', 'parentalidade', 'materno-infantil', 'infraestrutura escolar', 'política nacional de cuidados', 'Plano Nacional de Educação', 'Bolsa Família', 'Conanda', 'visitação domiciliar', 'Homeschooling', 'Política Nacional Integrada da Primeira Infância'],
    'IEPS': ['SUS', 'Sistema Único de Saúde', 'fortalecimento', 'Universalidade', 'Equidade em saúde', 'populações vulneráveis', 'desigualdades sociais', 'Organização do SUS', 'gestão pública', 'políticas públicas em saúde', 'Governança do SUS', 'regionalização', 'descentralização', 'Regionalização em saúde', 'Políticas públicas em saúde', 'População negra em saúde', 'Saúde indígena', 'Povos originários', 'Saúde da pessoa idosa', 'envelhecimento ativo', 'Atenção Primária', 'Saúde da criança', 'Saúde do adolescente', 'Saúde da mulher', 'Saúde do homem', 'Saúde da pessoa com deficiência', 'Saúde da população LGBTQIA+', 'Financiamento da saúde', 'atenção primária', 'tripartite', 'orçamento', 'Emendas e orçamento da saúde', 'Ministério da Saúde', 'Trabalhadores de saúde', 'Força de trabalho em saúde', 'Recursos humanos em saúde', 'Formação profissional de saúde', 'Cuidados primários em saúde', 'Emergências climáticas e ambientais em saúde', 'mudanças climáticas', 'adaptação climática', 'saúde ambiental', 'políticas climáticas', 'Vigilância em saúde', 'epidemiológica', 'Emergência em saúde', 'estado de emergência', 'Saúde suplementar', 'complementar', 'privada', 'planos de saúde', 'seguros', 'seguradoras', 'planos populares', 'Anvisa', 'gestão', 'governança', 'ANS', 'Sandbox regulatório', 'Cartões e administradoras de benefícios em saúde', 'Economia solidária em saúde mental', 'Pessoa em situação de rua', 'saúde mental', 'Fiscalização de comunidades terapêuticas', 'Rede de atenção psicossocial', 'RAPS', 'unidades de acolhimento', 'assistência multiprofissional', 'centros de convivência', 'Cannabis', 'canabidiol', 'tratamento terapêutico', 'Desinstitucionalização', 'manicômios', 'hospitais de custódia', 'Saúde mental na infância', 'adolescência', 'escolas', 'comunidades escolares', 'protagonismo juvenil', 'Dependência química', 'vícios', 'ludopatia', 'Treinamento em saúde mental', 'capacitação em saúde mental', 'Intervenções terapêuticas em saúde mental', 'Internet e redes sociais na saúde mental', 'Violência psicológica', 'Surto psicótico'],
    'Manual': ['Ozempic', 'Wegovy', 'Mounjaro', 'Telemedicina', 'Telessaúde', 'CBD', 'Cannabis Medicinal', 'CFM', 'Conselho Federal de Medicina', 'Farmácia Magistral', 'Medicamentos Manipulados', 'Minoxidil', 'Emagrecedores', 'Retenção de receita de medicamentos'],
    'Mevo': ['Prontuário eletrônico', 'dispensação eletrônica', 'telessaúde', 'assinatura digital', 'certificado digital', 'controle sanitário', 'prescrição por enfermeiros', 'doenças crônicas', 'autonomia da ANPD', 'Acesso e uso de dados', 'responsabilização de plataformas digitais', 'regulamentação de marketplaces', 'segurança cibernética', 'inteligência artificial', 'digitalização do SUS', 'venda de medicamentos', 'distribuição de medicamentos', 'Bula digital', 'Atesta CFM', 'SNGPC', 'Farmacêutico Remoto', 'Medicamentos Isentos de Prescrição', 'MIPs', 'RNDS', 'Rede Nacional de Dados em Saúde'],
    'Giro de notícias': ['Governo Lula', 'Presidente Lula', 'Governo', 'Governo Federal', 'Governo economia', 'Economia', 'Governo internacional', 'Saúde', 'Medicamento', 'Vacina', 'Câncer', 'Oncologia', 'Gripe', 'Diabetes', 'Obesidade', 'Alzheimer', 'Saúde mental', 'Síndrome respiratória', 'SUS', 'Sistema Único de Saúde', 'Ministério da Saúde', 'Alexandre Padilha', 'ANVISA', 'Primeira Infância', 'Infância', 'Criança', 'Saúde criança', 'Saúde infantil', 'cuidado criança', 'legislação criança', 'direitos da criança', 'criança câmara', 'criança senado', 'alfabetização', 'creche', 'ministério da educação', 'educação', 'educação Brasil', 'escolas', 'aprendizado', 'ensino integral', 'ensino médio', 'Camilo Santana'],
    'Cactus': ['Saúde mental', 'saúde mental para meninas', 'saúde mental para juventude', 'saúde mental para mulheres', 'Rede de atenção psicossocial', 'RAPS', 'CAPS', 'Centro de Apoio Psicossocial'],
    'Vital Strategies': ['Saúde mental', 'Dados para a saúde', 'Morte evitável', 'Doenças crônicas não transmissíveis', 'Rotulagem de bebidas alcoólicas', 'Educação em saúde', 'Bebidas alcoólicas', 'Imposto seletivo', 'Rotulagem de alimentos', 'Alimentos ultraprocessados', 'Publicidade infantil', 'Publicidade de alimentos ultraprocessados', 'Tributação de bebidas alcoólicas', 'Alíquota de bebidas alcoólicas', 'Cigarro eletrônico', 'Controle de tabaco', 'Violência doméstica', 'Exposição a fatores de risco', 'Departamento de Saúde Mental', 'Hipertensão arterial', 'Saúde digital', 'Violência contra crianças', 'Violência contra mulheres', 'Feminicídio', 'COP 30']
}

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

def _ensure_worksheets(spreadsheet, client_name: str):
    # Aba do cliente
    try:
        ws_data = spreadsheet.worksheet(client_name)
    except gspread.WorksheetNotFound:
        ws_data = spreadsheet.add_worksheet(title=client_name, rows="100", cols=str(len(HEADERS_DATA)))
        ws_data.update(values=[HEADERS_DATA], range_name=f'A1:{_col_letter(len(HEADERS_DATA))}1')
    # garante cabeçalho mesmo se a aba já existia sem header correto
    first_row = ws_data.row_values(1)
    if [c.strip().lower() for c in first_row] != [c.lower() for c in HEADERS_DATA]:
        ws_data.update(values=[HEADERS_DATA], range_name=f'A1:{_col_letter(len(HEADERS_DATA))}1')

    # FINGERPRINTS (continua sendo uma aba única)
    try:
        ws_fp = spreadsheet.worksheet(WORKSHEET_FP)
    except gspread.WorksheetNotFound:
        ws_fp = spreadsheet.add_worksheet(title=WORKSHEET_FP, rows="100", cols=str(len(HEADERS_FP)))
        ws_fp.update(values=[HEADERS_FP], range_name=f'A1:{_col_letter(len(HEADERS_FP))}1')
    first_row_fp = ws_fp.row_values(1)
    if [c.strip().lower() for c in first_row_fp] != [c.lower() for c in HEADERS_FP]:
        ws_fp.update(values=[HEADERS_FP], range_name=f'A1:{_col_letter(len(HEADERS_FP))}1')

    return ws_data, ws_fp

def load_fingerprints(gc) -> set:
    print("🧩 Carregando fingerprints do Google Sheets...")
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    try:
        ws_fp = spreadsheet.worksheet(WORKSHEET_FP)
        vals = ws_fp.get_all_records()
        fps = {row.get('fp') for row in vals if row.get('fp')}
        print(f"🧩 Fingerprints carregados: {len(fps)}")
        return fps
    except gspread.WorksheetNotFound:
        print("🧩 Aba de fingerprints não encontrada. Será criada ao inserir novos dados.")
        return set()

def insert_fingerprints(gc, new_fps: list[str]):
    if not new_fps:
        return
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    try:
        ws_fp = spreadsheet.worksheet(WORKSHEET_FP)
    except gspread.WorksheetNotFound:
        # Cria a aba se não existir
        ws_fp = spreadsheet.add_worksheet(title=WORKSHEET_FP, rows="100", cols=str(len(HEADERS_FP)))
        ws_fp.update(values=[HEADERS_FP], range_name=f'A1:{_col_letter(len(HEADERS_FP))}1')
    
    now_brt = fmt_brasilia(agora_brasilia())
    rows = [[fp, now_brt] for fp in new_fps]
    # Insere fingerprints logo após o cabeçalho
    ws_fp.insert_rows(rows, row=2, value_input_option='RAW')
    print(f"🧩 Fingerprints inseridos: {len(new_fps)}")

def insert_news_rows_by_client(gc, df_por_cliente: dict[str, pd.DataFrame]):
    """Insert na aba de cada cliente."""
    spreadsheet = gc.open_by_key(SPREADSHEET_KEY)
    
    for client_name, df_client in df_por_cliente.items():
        if df_client.empty:
            print(f"ℹ️ [{client_name}] Nada novo para adicionar.")
            continue
            
        ws_data, _ = _ensure_worksheets(spreadsheet, client_name)

        # normaliza/ordena colunas
        for col in HEADERS_DATA:
            if col not in df_client.columns:
                df_client[col] = ''
        df_client = df_client[HEADERS_DATA].fillna('')

        # Insere linhas a partir da linha 2 (logo após o cabeçalho)
        ws_data.insert_rows(df_client.values.tolist(), row=2, value_input_option='USER_ENTERED')
        print(f"✅ [{client_name}] Linhas inseridas no início da aba: {len(df_client)}")

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
                    'Data de Publicação': dt_utc.astimezone(TZ_BR).strftime('%Y-%m-%d') if dt_utc else 'Data não encontrada',
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
# Main (com abas por cliente)
# =========================
def main():
    print("🚀 Iniciando scraper de notícias por cliente...")
    driver = setup_driver()

    # Dicionário para armazenar resultados por cliente
    resultados_por_cliente = {client: [] for client in CLIENT_SEARCH_TERMS.keys()}
    resumo_coletas = []

    try:
        # Processar termos de busca por cliente
        for client, search_terms in CLIENT_SEARCH_TERMS.items():
            print(f"\n📌 Processando cliente: {client}")
            for termo in tqdm(search_terms, desc=f"🔎 [{client}] Buscando termos"):
                noticias = scrape_news_for_term(driver, termo)
                df_noticias = pd.DataFrame(noticias)
                df_noticias_24h = filter_recent_24h(df_noticias)
                resultados_por_cliente[client].append(df_noticias_24h)
                resumo_coletas.append({
                    'Cliente': client,
                    'Termo de Busca': termo, 
                    'Notícias Coletadas (24h)': len(df_noticias_24h)
                })

        # Consolidar resultados por cliente
        df_por_cliente = {}
        for client, lista_dfs in resultados_por_cliente.items():
            if lista_dfs:
                df_client = pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()
                
                # Dedup local por URL Final (ou Link) dentro do cliente
                if not df_client.empty:
                    if 'URL Final' in df_client.columns:
                        df_client = df_client.drop_duplicates(subset=['URL Final']).copy()
                    else:
                        df_client = df_client.drop_duplicates(subset=['Link']).copy()
                    
                    # Adicionar timestamp de coleta
                    df_client['Coletado em'] = fmt_brasilia(agora_brasilia())
                
                df_por_cliente[client] = df_client
            else:
                df_por_cliente[client] = pd.DataFrame()

        # Carregar fingerprints existentes
        new_fps = []
        df_final_por_cliente = {}
        
        try:
            gc = _gspread_client_from_env()
            _ = gc.open_by_key(SPREADSHEET_KEY)  # sanity check / permissão
        except Exception as e:
            print(f"❌ Erro com credenciais/planilha: {e}")
            sys.exit(1)

        # Dedup remota via fingerprints (global entre todos os clientes)
        existing = load_fingerprints(gc)
        
        for client, df_client in df_por_cliente.items():
            if not df_client.empty:
                df_client['__fp'] = df_client.apply(make_fingerprint, axis=1)
                mask_novos = ~df_client['__fp'].isin(existing)
                df_novo = df_client[mask_novos].copy()
                
                # Adicionar novos fingerprints ao conjunto
                new_client_fps = df_novo['__fp'].tolist()
                new_fps.extend(new_client_fps)
                existing.update(new_client_fps)  # Atualiza o conjunto local para evitar duplicatas entre clientes
                
                if '__fp' in df_novo.columns:
                    df_novo.drop(columns=['__fp'], inplace=True)
                
                df_final_por_cliente[client] = df_novo
            else:
                df_final_por_cliente[client] = pd.DataFrame()

        # INSERT no Google Sheets (por cliente)
        sheets_ok = False
        try:
            if any(not df.empty for df in df_final_por_cliente.values()):
                insert_news_rows_by_client(gc, df_final_por_cliente)
            else:
                print("ℹ️ Nenhuma notícia nova após deduplicação por fingerprints.")
            
            if new_fps:
                insert_fingerprints(gc, new_fps)
            
            sheets_ok = True
        except Exception as e:
            print(f"❌ Erro ao fazer insert no Google Sheets: {e}")
            sheets_ok = False

        # Resumo final
        print(f"\n📊 RESUMO DA EXECUÇÃO:")
        print(f"🕒 Coletado em (BRT): {fmt_brasilia(agora_brasilia())}")
        print(f"📤 Google Sheets (insert): {'✅' if sheets_ok else '❌'}")
        
        # Resumo por cliente
        for client, df_novo in df_final_por_cliente.items():
            total_noticias = len(df_novo) if not df_novo.empty else 0
            print(f"📰 [{client}] Total de notícias novas: {total_noticias}")
        
        # Resumo detalhado
        df_resumo = pd.DataFrame(resumo_coletas)
        if not df_resumo.empty:
            print("\n📋 Detalhamento por termo:")
            for _, row in df_resumo.iterrows():
                if row['Notícias Coletadas (24h)'] > 0:
                    print(f"   [{row['Cliente']}] {row['Termo de Busca']}: {row['Notícias Coletadas (24h)']} notícias")

    except Exception as e:
        print(f"❌ Erro geral na execução: {e}")
        sys.exit(1)
    finally:
        driver.quit()
        print("🏁 Scraper finalizado")

if __name__ == "__main__":
    main()
