# Alinhamento baseado em título+resumo+texto_completo (todas as abas)
# - Usa Google Sheets (service account) e Gemini 2.5 Flash
# - Lê todas as abas; para cada linha com material e sem "Alinhamento", escreve Alinhamento e Justificativa
# - Robustez: backoff 429, valida JSON, tamanho máx. de célula

import os, time, json, re, random, html
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
from google import genai
from string import Template

# Config via env
GENAI_API_KEY = os.getenv("GENAI_API_KEY", "").strip()
assert GENAI_API_KEY, "Defina o secret GENAI_API_KEY."
MODEL_NAME = os.getenv("GENAI_MODEL", "gemini-2.5-flash").strip()

PLANILHA = os.getenv("PLANILHA", "").strip()
assert PLANILHA, "Defina o secret PLANILHA (ID da planilha do Google Sheets)."

CREDENTIALS_JSON = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")

# Colunas de entrada (case-insensitive)
TIT_COL  = os.getenv("ALIGN_COL_TITULO", "titulo")
RES_COL  = os.getenv("ALIGN_COL_RESUMO", "resumo")
BODY_COL = os.getenv("ALIGN_COL_TEXTO", "texto_completo")

# Colunas de saída
OUT_ALINH_COL = os.getenv("ALIGN_COL_SAIDA1", "Alinhamento")
OUT_JUST_COL  = os.getenv("ALIGN_COL_SAIDA2", "Justificativa")

# Ajustes operacionais
BATCH_SIZE = int(os.getenv("ALIGN_BATCH_SIZE", "20"))
SLEEP_SEC  = float(os.getenv("ALIGN_SLEEP_SEC", "0"))
READ_RANGE = os.getenv("ALIGN_READ_RANGE", "")
SKIP_TITLES = [s.strip() for s in os.getenv("ALIGN_SKIP_TABS", "").split(",") if s.strip()]

# Segurança p/ células do Sheets (limite ~50k)
MAX_CELL_CHARS = int(os.getenv("MAX_CELL_CHARS", "47000"))

# Mapa cliente → (nome, descrição) p/ personalizar análise por aba
ORG_MAP = {
    "IU": ("Instituto Unibanco (IU)",
           "O Instituto Unibanco (IU) apoia redes estaduais de ensino na melhoria da gestão educacional por meio de projetos, produção de conhecimento e apoio técnico."),
    "FMCSV": ("Fundação Maria Cecilia Souto Vidigal (FMCSV)",
              "Atua pela causa da primeira infância, conectando pesquisa, advocacy e apoio a políticas públicas para o desenvolvimento integral de crianças de 0 a 6 anos."),
    "IEPS": ("Instituto de Estudos para Políticas de Saúde (IEPS)",
             "Organização independente dedicada a aprimorar políticas de saúde no Brasil, com foco em atenção primária, saúde digital e financiamento do SUS."),
    "IAS": ("Instituto Ayrton Senna (IAS)",
            "Centro de inovação em educação que atua com aprendizagem acadêmica e competências socioemocionais."),
    "ISG": ("Instituto Sonho Grande (ISG)",
            "Apoia a expansão e qualificação do ensino médio integral em redes públicas."),
    "Reúna": ("Instituto Reúna",
              "Ferramentas e pesquisas para implementação de políticas educacionais alinhadas à BNCC."),
    "Reuna": ("Instituto Reúna",
              "Ferramentas e pesquisas para implementação de políticas educacionais alinhadas à BNCC."),
    "REMS": ("REMS – Rede Esporte pela Mudança Social",
             "Articula organizações que usam o esporte como vetor de desenvolvimento humano."),
    "Manual": ("Manual (saúde)",
               "Plataforma digital voltada à saúde masculina, com atendimento online e tratamentos baseados em evidências."),
    "Cactus": ("Instituto Cactus",
               "Atuação independente em saúde mental, priorizando adolescentes e mulheres, via advocacy e fomento."),
    "Vital Strategies": ("Vital Strategies",
                         "Organização global de saúde pública que apoia políticas baseadas em evidências."),
    "Mevo": ("Mevo",
             "Healthtech que integra soluções de saúde digital do consultório à entrega de medicamentos."),
    "Coletivo Feminista": ("Coletivo Feminista",
             "Movimento que atua por direitos reprodutivos e descriminalização do aborto no Brasil."),
}

# Prompt
PROMPT = Template("""
Você é analista de políticas públicas. Avalie a coerência do material abaixo com a missão do(a) $cliente.

Missão e escopo do cliente:
$cliente_descricao

Instruções:
- Baseie-se exclusivamente no material fornecido (compilado de título, resumo e texto completo).
- Classifique o alinhamento em um dos três valores: "Alinha", "Parcial" ou "Não Alinha".
- Escreva uma justificativa breve (1 a 3 frases), objetiva, citando elementos do material.
- Responda somente em JSON válido, sem comentários.

Formato:
{"alinhamento":"Alinha|Parcial|Não Alinha","justificativa":"texto"}

Material:
\"\"\"$material\"\"\"
""".strip())

# Conexões
genai_client = genai.Client(api_key=GENAI_API_KEY)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(CREDENTIALS_JSON, scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(PLANILHA)

# Helpers
def norm_find_col(df, name):
    target = name.strip().lower()
    for c in df.columns:
        if c.strip().lower() == target:
            return c
    return None

def strip_html(text: str) -> str:
    if not isinstance(text, str): return ""
    t = html.unescape(text)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def _truncate_cell(s: str, limit: int = MAX_CELL_CHARS) -> str:
    s = str(s) if s is not None else ""
    return s if len(s) <= limit else s[: max(0, limit - 10)] + " [..]"

def build_material(row, tit_col, res_col, body_col) -> str:
    t = strip_html(row.get(tit_col, ""))
    r = strip_html(row.get(res_col, ""))
    b = strip_html(row.get(body_col, ""))
    parts = []
    if t: parts.append(f"Título: {t}")
    if r: parts.append(f"Resumo: {r}")
    if b: parts.append(f"Texto: {b}")
    material = "\n\n".join(parts).strip()
    if len(material) > 12000:
        material = material[:11990] + " [..]"
    return material

def read_sheet_df(ws, read_range: str = "") -> pd.DataFrame:
    def _once():
        vals = ws.get(read_range) if read_range else ws.get_all_values()
        if not vals: return pd.DataFrame()
        header, data = vals[0], vals[1:]
        width = len(header)
        data = [row + [""] * (width - len(row)) for row in data]
        return pd.DataFrame(data, columns=[h.strip() for h in header])
    delay = 1.0
    for _ in range(6):
        try:
            return _once()
        except gspread.exceptions.APIError as e:
            msg = str(e)
            if "429" in msg or "Quota exceeded" in msg:
                time.sleep(delay + random.random()*0.25)
                delay = min(delay*2, 20)
                continue
            raise
    return _once()

def call_gemini(prompt_text: str) -> dict:
    delay = 1.0
    for _ in range(5):
        try:
            stream = genai_client.models.generate_content_stream(
                model=MODEL_NAME,
                contents=prompt_text,
                config={"response_mime_type": "application/json"},
            )
            raw = "".join((chunk.text or "") for chunk in stream).strip()
            m = re.search(r"\{.*\}", raw, flags=re.S)
            if not m:
                return {"alinhamento": "Parcial", "justificativa": "Saída sem JSON; revisar."}
            data = json.loads(m.group(0))
            alinh = str(data.get("alinhamento", "")).strip() or "Parcial"
            just  = str(data.get("justificativa", "")).strip() or "Sem justificativa; revisar."
            if alinh not in ("Alinha", "Parcial", "Não Alinha"):
                alinh = "Parcial"
            return {"alinhamento": alinh, "justificativa": just}
        except Exception:
            time.sleep(delay + random.random()*0.25)
            delay = min(delay*2, 20)
    return {"alinhamento": "Parcial", "justificativa": "Falha após tentativas; revisar."}

def classify_material(material: str, nome_cli: str, desc_cli: str) -> dict:
    if not material:
        return {"alinhamento": "Parcial",
                "justificativa": "Sem conteúdo em título, resumo ou texto; não é possível concluir o alinhamento."}
    prompt_text = PROMPT.substitute(cliente=nome_cli, cliente_descricao=desc_cli, material=material)
    return call_gemini(prompt_text)

def process_sheet(ws):
    title = ws.title.strip()
    if title in SKIP_TITLES:
        print(f"⏭️  Pulando aba '{title}'."); return

    nome_cli, desc_cli = ORG_MAP.get(title, (title, ""))

    print(f"\n▶️  Aba: {title} | Cliente: {nome_cli}")
    df = read_sheet_df(ws, READ_RANGE)
    if df.empty:
        print(f"[{title}] vazia — pulando."); return

    df.columns = [c.strip() for c in df.columns]

    tit_col  = norm_find_col(df, TIT_COL)
    res_col  = norm_find_col(df, RES_COL)
    body_col = norm_find_col(df, BODY_COL)

    if not any([tit_col, res_col, body_col]):
        print(f"[{title}] não achei '{TIT_COL}', '{RES_COL}' ou '{BODY_COL}' — pulando.")
        return

    if OUT_ALINH_COL not in df.columns:
        df[OUT_ALINH_COL] = ""
    if OUT_JUST_COL not in df.columns:
        df[OUT_JUST_COL] = ""

    def row_has_material(i):
        cols = [c for c in [tit_col, res_col, body_col] if c]
        return any(str(df.at[i, c]).strip() for c in cols)

    to_process = [i for i in range(len(df))
                  if not str(df.at[i, OUT_ALINH_COL]).strip() and row_has_material(i)]

    print(f"[{title}] linhas para classificar: {len(to_process)}")
    if not to_process:
        return

    for start in range(0, len(to_process), BATCH_SIZE):
        batch_idx = to_process[start:start+BATCH_SIZE]
        for i in batch_idx:
            material = build_material(df.loc[i], tit_col, res_col, body_col)
            res = classify_material(material, nome_cli, desc_cli)
            df.at[i, OUT_ALINH_COL] = _truncate_cell(res["alinhamento"], MAX_CELL_CHARS)
            df.at[i, OUT_JUST_COL]  = _truncate_cell(res["justificativa"], MAX_CELL_CHARS)
            if SLEEP_SEC:
                time.sleep(SLEEP_SEC)

        set_with_dataframe(ws, df.iloc[:max(batch_idx)+1],
                           include_index=False, include_column_header=True, resize=False)
        print(f"[{title}] 💾 salvo linhas até {max(batch_idx)+2}")

def main():
    worksheets = sh.worksheets()
    if not worksheets:
        print("Planilha sem abas."); return
    for ws in worksheets:
        process_sheet(ws)
    print("\n✅ Concluído (todas as abas).")

if __name__ == "__main__":
    main()
