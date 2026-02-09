import os, time, json, re, random, html
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
from google import genai
from string import Template

GENAI_API_KEY = os.getenv("GENAI_API_KEY", "").strip()
assert GENAI_API_KEY, "Defina o secret GENAI_API_KEY."
MODEL_NAME = os.getenv("GENAI_MODEL", "gemini-2.5-flash").strip()

PLANILHA = os.getenv("PLANILHA", "").strip()
assert PLANILHA, "Defina o secret PLANILHA (ID da planilha do Google Sheets)."

CREDENTIALS_JSON = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")

TIT_COL  = os.getenv("ALIGN_COL_TITULO", "titulo")
RES_COL  = os.getenv("ALIGN_COL_RESUMO", "resumo")
BODY_COL = os.getenv("ALIGN_COL_TEXTO", "texto_completo")

OUT_ALINH_COL = os.getenv("ALIGN_COL_SAIDA1", "Alinhamento")
OUT_JUST_COL  = os.getenv("ALIGN_COL_SAIDA2", "Justificativa")

BATCH_SIZE = int(os.getenv("ALIGN_BATCH_SIZE", "20"))
SLEEP_SEC  = float(os.getenv("ALIGN_SLEEP_SEC", "0"))
READ_RANGE = os.getenv("ALIGN_READ_RANGE", "")
SKIP_TITLES = [s.strip() for s in os.getenv("ALIGN_SKIP_TABS", "").split(",") if s.strip()]

MAX_CELL_CHARS = int(os.getenv("MAX_CELL_CHARS", "47000"))

# Liga a remoção automática (tira Não Alinha e Não se aplica)
DELETE_INVALID = os.getenv("DELETE_INVALID", "1").strip() in ("1","true","True","yes","on")
DELETE_CHUNK_SIZE = int(os.getenv("DELETE_CHUNK_SIZE", "80"))

ORG_MAP = {
    "IU": ("Instituto Unibanco (IU)",
           "O Instituto Unibanco (IU) é uma organização sem fins lucrativos que apoia redes estaduais de ensino na melhoria da gestão educacional por meio de projetos como o Jovem de Futuro, produção de conhecimento e apoio técnico a secretarias de educação."),
    "FMCSV": ("Fundação Maria Cecilia Souto Vidigal (FMCSV)",
              "A Fundação Maria Cecilia Souto Vidigal (FMCSV) atua pela causa da primeira infância no Brasil, conectando pesquisa, advocacy e apoio a políticas públicas para garantir o desenvolvimento integral de crianças de 0 a 6 anos; iniciativas como o “Primeira Infância Primeiro” oferecem dados e ferramentas para gestores e candidatos."),
    "IEPS": ("Instituto de Estudos para Políticas de Saúde (IEPS)",
             "O Instituto de Estudos para Políticas de Saúde (IEPS) é uma organização independente e sem fins lucrativos dedicada a aprimorar políticas de saúde no Brasil, combinando pesquisa aplicada, produção de evidências e advocacy em temas como atenção primária, saúde digital e financiamento do SUS."),
    "IAS": ("Instituto Ayrton Senna (IAS)",
            "O Instituto Ayrton Senna (IAS) é um centro de inovação em educação que atua em pesquisa e desenvolvimento, disseminação em larga escala e influência em políticas públicas, com foco em aprendizagem acadêmica e competências socioemocionais na rede pública."),
    "ISG": ("Instituto Sonho Grande (ISG)",
            "O Instituto Sonho Grande (ISG) é uma organização sem fins lucrativos e apartidária voltada à expansão e qualificação do ensino médio integral em redes públicas; trabalha em parceria com estados para revisão curricular, formação de equipes e gestão orientada a resultados."),
    "Reúna": ("Instituto Reúna",
              "O Instituto Reúna desenvolve pesquisas e ferramentas para apoiar redes e escolas na implementação de políticas educacionais alinhadas à BNCC, com foco em currículo, materiais de apoio e formação de professores."),
    "Reuna": ("Instituto Reúna",
              "O Instituto Reúna desenvolve pesquisas e ferramentas para apoiar redes e escolas na implementação de políticas educacionais alinhadas à BNCC, com foco em currículo, materiais de apoio e formação de professores."),
    "REMS": ("REMS – Rede Esporte pela Mudança Social",
             "A REMS – Rede Esporte pela Mudança Social articula organizações que usam o esporte como vetor de desenvolvimento humano, mobilizando atores e produzindo conhecimento para ampliar o impacto social dessa agenda no país."),
    "Manual": ("Manual (saúde)",
               "A Manual (saúde) é uma plataforma digital voltada principalmente à saúde masculina, oferecendo atendimento online e tratamentos baseados em evidências (como saúde capilar, sono e saúde sexual), com prescrição médica e acompanhamento remoto."),
    "Cactus": ("Instituto Cactus",
               "O Instituto Cactus é uma entidade filantrópica e de direitos humanos que atua de forma independente em saúde mental, priorizando adolescentes e mulheres, por meio de advocacy e fomento a projetos de prevenção e promoção de cuidado em saúde mental."),
    "Vital Strategies": ("Vital Strategies",
                         "A Vital Strategies é uma organização global de saúde pública que trabalha com governos e sociedade civil na concepção e implementação de políticas baseadas em evidências em áreas como doenças crônicas, segurança viária, qualidade do ar, dados vitais e comunicação de risco."),
    "Mevo": ("Mevo",
             "A Mevo é uma healthtech brasileira que integra soluções de saúde digital (da prescrição eletrônica à compra/entrega de medicamentos) conectando médicos, hospitais, farmácias e pacientes para tornar o cuidado mais simples e rastreável."),
    "Coletivo Feminista": ("Coletivo Feminista",
                          "O Coletivo Feminista é um movimento feminista que atua pela descriminalização e legalização do aborto no Brasil, articulando pesquisa, incidência política e mobilização social. Seus princípios ético-políticos abrangem a comunicação como direito e fundamento da democracia, a defesa do Estado democrático de direito, a compreensão de que maternidade não é dever e deve respeitar a liberdade de escolha, a promoção de uma atenção universal, equânime e integral à saúde — com ênfase no papel do SUS, no acesso a métodos contraceptivos e abortivos seguros e no respeito à autodeterminação reprodutiva —, além da defesa da descriminalização e legalização do aborto."),
    "IDEC": ("Instituto Brasileiro de Defesa do Consumidor (Idec)",
             "O Instituto Brasileiro de Defesa do Consumidor (Idec) é uma associação civil sem fins lucrativos e independente de empresas, partidos ou governos, fundada em 1987. Atua na defesa dos direitos dos consumidores e na promoção de relações de consumo éticas, seguras e sustentáveis. Sua agenda combina advocacy, pesquisa e litigância estratégica, com foco em temas como saúde, alimentação, energia, telecomunicações e proteção de dados pessoais. O Idec se destaca na promoção de políticas públicas voltadas à alimentação saudável, ao controle de ultraprocessados e agrotóxicos, à rotulagem nutricional, à transição energética justa e à regulação de plataformas digitais."),
    "Umane": ("Umane",
              "A Umane é uma organização da sociedade civil, isenta e sem fins lucrativos, que atua para fomentar melhorias sistêmicas na saúde pública no Brasil, apoiando iniciativas baseadas em evidências para ampliar equidade, eficiência e qualidade do sistema. Trabalha com fomento a projetos, articulação com parceiros e monitoramento e avaliação, com frentes como Atenção Primária à Saúde (APS), Doenças Crônicas Não Transmissíveis (DCNT) e saúde da mulher, da criança e do adolescente.")
}

PROMPT = Template(r"""
Você é analista de políticas públicas e faz triagem de atos do DOU, matérias legislativas e notícias para um(a) cliente.

Missão/escopo do cliente:
$descricao

Tarefa:
Classificar o alinhamento do **Conteúdo** com a missão do cliente.

Regras de evidência:
- Use **apenas** o Conteúdo. Não use contexto externo.
- NÃO exija que o Conteúdo cubra TODA a missão do cliente.
  # Se o Conteúdo estiver claramente dentro de ao menos UMA frente/eixo relevante do cliente, marque "Alinha".
  # A ausência de menção a outras frentes (ex.: socioemocional) NÃO reduz automaticamente para "Parcial".
- Use "Parcial" apenas quando houver INSUFICIÊNCIA ou AMBIGUIDADE no texto para decidir.
- Se o texto for claramente de natureza incompatível com triagem temática (ex.: decisão sobre caso individual sem política pública; deferimento/indeferimento nominal; concessão pontual; nomeação/dispensa rotineira sem tema; mero expediente administrativo sem objeto; publicação que não permite inferir assunto), marque "Não se aplica".

Classes (escolha exatamente UMA):
- "Alinha": O objeto/tema do Conteúdo é claro e há evidência explícita de relação com pelo menos 1 frente/eixo do cliente.
- "Parcial": O Conteúdo sugere relação, mas é genérico, incompleto ou não permite identificar com segurança o objeto/tema.
- "Não Alinha": O tema é claro e não tem relação com a missão do cliente.
- "Não se aplica": O Conteúdo não é classificável por tema/escopo do cliente com base no texto (exemplos acima), ou é predominantemente ato individual/procedimental sem política pública inferível.

Formato de saída:
Retorne **somente** JSON válido neste formato:
{
  "alinhamento": "Alinha" | "Parcial" | "Não Alinha" | "Não se aplica",
  "justificativa": "1–3 frases citando elementos do Conteúdo (termos/trechos) que sustentam a decisão"
}

Conteúdo:
\"\"\"$conteudo\"\"\"
""".strip())

genai_client = genai.Client(api_key=GENAI_API_KEY)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(CREDENTIALS_JSON, scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(PLANILHA)

def norm_find_col(df, name):
    target = name.strip().lower()
    for c in df.columns:
        if c.strip().lower() == target:
            return c
    return None

def strip_html(text: str) -> str:
    if not isinstance(text, str):
        return ""
    t = html.unescape(text)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def _truncate_cell(s: str, limit: int = MAX_CELL_CHARS) -> str:
    s = str(s) if s is not None else ""
    return s if len(s) <= limit else s[: max(0, limit - 10)] + " [..]"

def build_content(row, tit_col, res_col, body_col) -> str:
    t = strip_html(row.get(tit_col, ""))
    r = strip_html(row.get(res_col, ""))
    b = strip_html(row.get(body_col, ""))
    parts = []
    if t: parts.append(f"Título: {t}")
    if r: parts.append(f"Resumo: {r}")
    if b: parts.append(f"Texto: {b}")
    conteudo = "\n\n".join(parts).strip()
    if len(conteudo) > 12000:
        conteudo = conteudo[:11990] + " [..]"
    return conteudo

def read_sheet_df(ws, read_range: str = "") -> pd.DataFrame:
    def _once():
        vals = ws.get(read_range) if read_range else ws.get_all_values()
        if not vals:
            return pd.DataFrame()
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
                time.sleep(delay + random.random() * 0.25)
                delay = min(delay * 2, 20)
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
            if alinh not in ("Alinha", "Parcial", "Não Alinha", "Não se aplica"):
                alinh = "Parcial"
            return {"alinhamento": alinh, "justificativa": just}
        except Exception:
            time.sleep(delay + random.random() * 0.25)
            delay = min(delay * 2, 20)
    return {"alinhamento": "Parcial", "justificativa": "Falha após tentativas; revisar."}

def classify_content(conteudo: str, descricao_cliente: str) -> dict:
    if not conteudo:
        return {
            "alinhamento": "Parcial",
            "justificativa": "Sem conteúdo em título, resumo ou texto; não é possível concluir o alinhamento."
        }
    prompt_text = PROMPT.substitute(descricao=descricao_cliente, conteudo=conteudo)
    return call_gemini(prompt_text)

def _range_start_row(read_range: str) -> int:
    if not read_range:
        return 1
    m = re.match(r"^\s*[A-Za-z]+\s*(\d+)", read_range)
    if m:
        return int(m.group(1))
    m2 = re.match(r"^\s*[A-Za-z]+\s*(\d+)\s*:", read_range)
    if m2:
        return int(m2.group(1))
    return 1

def _delete_rows_in_chunks(ws, rows_1based, chunk_size=80):
    rows = sorted(set(int(r) for r in rows_1based if int(r) >= 2), reverse=True)
    if not rows:
        return 0
    deleted = 0
    for start in range(0, len(rows), chunk_size):
        chunk = rows[start:start + chunk_size]
        for r in chunk:
            ws.delete_rows(r)
            deleted += 1
        time.sleep(0.2)
    return deleted

def _is_invalid_alignment(v):
    s = str(v).strip().lower()
    return s in (
        "não se aplica", "nao se aplica", "não seaplica", "nao seaplica",
        "nao-se-aplica", "não-se-aplica"
    )

def process_sheet(ws):
    title = ws.title.strip()
    if title in SKIP_TITLES:
        print(f"⏭️  Pulando aba '{title}'.")
        return

    _, desc_cli = ORG_MAP.get(title, (title, ""))

    print(f"\n▶️  Aba: {title}")
    df = read_sheet_df(ws, READ_RANGE)
    if df.empty:
        print(f"[{title}] vazia — pulando.")
        return

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

    to_process = [
        i for i in range(len(df))
        if not str(df.at[i, OUT_ALINH_COL]).strip() and row_has_material(i)
    ]

    print(f"[{title}] linhas para classificar: {len(to_process)}")
    if to_process:
        for start in range(0, len(to_process), BATCH_SIZE):
            batch_idx = to_process[start:start + BATCH_SIZE]
            for i in batch_idx:
                conteudo = build_content(df.loc[i], tit_col, res_col, body_col)
                res = classify_content(conteudo, desc_cli)
                df.at[i, OUT_ALINH_COL] = _truncate_cell(res["alinhamento"], MAX_CELL_CHARS)
                df.at[i, OUT_JUST_COL]  = _truncate_cell(res["justificativa"], MAX_CELL_CHARS)
                if SLEEP_SEC:
                    time.sleep(SLEEP_SEC)

            set_with_dataframe(
                ws,
                df.iloc[:max(batch_idx) + 1],
                include_index=False,
                include_column_header=True,
                resize=False
            )
            print(f"[{title}] 💾 salvo linhas até {max(batch_idx) + 2}")

    if not DELETE_INVALID:
        return

    print(f"[{title}] 🧹 removendo linhas com 'Não Alinha' e 'Não se aplica'...")
    start_row = _range_start_row(READ_RANGE)
    data_start_row = start_row + 1

    col_alinh = norm_find_col(df, OUT_ALINH_COL) or OUT_ALINH_COL
    if col_alinh not in df.columns:
        print(f"[{title}] não existe coluna '{OUT_ALINH_COL}' — nada a remover.")
        return

    idx_to_drop = [i for i in range(len(df)) if _is_invalid_alignment(df.at[i, col_alinh])]
    if not idx_to_drop:
        print(f"[{title}] nada para remover.")
        return

    sheet_rows_to_delete = [data_start_row + i for i in idx_to_drop]
    deleted = _delete_rows_in_chunks(ws, sheet_rows_to_delete, chunk_size=DELETE_CHUNK_SIZE)
    print(f"[{title}] ✅ removidas {deleted} linhas.")

def main():
    worksheets = sh.worksheets()
    if not worksheets:
        print("Planilha sem abas.")
        return
    for ws in worksheets:
        process_sheet(ws)
    print("\n✅ Concluído (todas as abas).")

if __name__ == "__main__":
    main()
