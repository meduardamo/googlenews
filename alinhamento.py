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

DELETE_INVALID = os.getenv("DELETE_INVALID", "1").strip() in ("1","true","True","yes","on")
DELETE_CHUNK_SIZE = int(os.getenv("DELETE_CHUNK_SIZE", "80"))

CLIENTE_DESCRICOES = {
    "IU": ("Instituto Unibanco (IU)",
           "O Instituto Unibanco (IU) é uma organização sem fins lucrativos que atua no fortalecimento da gestão educacional, desenvolvendo projetos como o Jovem de Futuro, oferecendo apoio técnico a secretarias estaduais de educação e produzindo conhecimento para aprimorar políticas públicas. Seu foco está tanto no cenário federal, acompanhando os debates sobre o financiamento da educação, programas nacionais de educação, regulação educacional e diretrizes definidas por órgãos como o Conselho Nacional de Educação, quanto subnacional, olhando para 6 estados prioritários (RS, MG, ES, CE, PI e GO). O IU apoia iniciativas de recomposição de aprendizagens, infraestrutura escolar, inclusão digital, educação ambiental, mudanças do clima e valorização de profissionais da educação."),
    "FMCSV": ("Fundação Maria Cecilia Souto Vidigal (FMCSV)",
              "A Fundação Maria Cecilia Souto Vidigal (FMCSV) é uma organização da sociedade civil dedicada ao fortalecimento da primeira infância no Brasil. Sua atuação concentra-se na integração entre produção de conhecimento, advocacy e apoio à formulação e implementação de políticas públicas, com o objetivo de assegurar o desenvolvimento integral de crianças de 0 a 6 anos. A Fundação acompanha o debate sobre educação domiciliar (homeschooling), posicionando-se de forma contrária a avanços nessa pauta. Além disso, participa ativamente da construção e implementação da Política Nacional Integrada da Primeira Infância. Entre suas iniciativas, destaca-se o programa 'Primeira Infância Primeiro', que disponibiliza dados, evidências e ferramentas para gestores públicos e candidatos, contribuindo para a qualificação do debate e das políticas voltadas à infância."),
    "IEPS": ("Instituto de Estudos para Políticas de Saúde (IEPS)",
             "O Instituto de Estudos para Políticas de Saúde (IEPS) é uma organização independente e sem fins lucrativos dedicada a aprimorar políticas de saúde no Brasil, combinando pesquisa aplicada, produção de evidências e advocacy em temas como atenção primária, saúde digital e financiamento do SUS. Com especialização em políticas públicas de saúde, o IEPS possui uma atuação centrada no fortalecimento do SUS enquanto sistema universal e equitativo. Seus núcleos de observação abrangem: a organização e governança federativa do SUS (modelo tripartite, emendas parlamentares); estruturação da APS, financiamento per capita e regionalização; equidade e enfrentamento de desigualdades; força de trabalho em saúde; saúde mental e Rede de Atenção Psicossocial; regulação da Anvisa e ANS; saúde suplementar; emergências sanitárias; e impactos das mudanças climáticas na saúde. As BETs são destaque no monitoramento para 2026, especialmente na incidência junto à Frente Parlamentar Mista para a Promoção da Saúde Mental."),
    "IAS": ("Instituto Ayrton Senna (IAS)",
            "O Instituto Ayrton Senna (IAS) é um centro de inovação em educação que atua em pesquisa e desenvolvimento, disseminação em larga escala e influência em políticas públicas, com foco em aprendizagem acadêmica e competências socioemocionais na rede pública."),
    "ISG": ("Instituto Sonho Grande (ISG)",
            "O Instituto Sonho Grande (ISG) é uma organização sem fins lucrativos e apartidária voltada à expansão e qualificação do ensino médio integral em redes públicas. Atua em parceria com secretarias estaduais de educação, oferecendo apoio na revisão curricular, na formação de equipes escolares e na implementação de modelos de gestão orientados a resultados. Além disso, o Instituto mantém foco no fortalecimento da infraestrutura das escolas públicas de tempo integral e na promoção de políticas voltadas à alfabetização. Também acompanha o debate sobre educação domiciliar (homeschooling), posicionando-se de forma contrária à sua ampliação."),
    "Reúna": ("Instituto Reúna",
              "O Instituto Reúna desenvolve pesquisas e ferramentas para apoiar redes e escolas na implementação de políticas educacionais alinhadas à BNCC, com foco em currículo, materiais de apoio e formação de professores."),
    "Reuna": ("Instituto Reúna",
              "O Instituto Reúna desenvolve pesquisas e ferramentas para apoiar redes e escolas na implementação de políticas educacionais alinhadas à BNCC, com foco em currículo, materiais de apoio e formação de professores."),
    "REMS": ("REMS – Rede Esporte pela Mudança Social",
             "A REMS – Rede Esporte pela Mudança Social articula organizações que usam o esporte como vetor de desenvolvimento humano, mobilizando atores e produzindo conhecimento para ampliar o impacto social dessa agenda no país. A rede atua desde 2007 no fortalecimento do campo do esporte para o desenvolvimento social, promovendo a troca de experiências, a sistematização de práticas e a realização de agendas coletivas, bem como acompanhando e incidindo em debates sobre políticas públicas, financiamento, marcos regulatórios e programas governamentais relacionados ao esporte educacional, comunitário e de participação. Sua atuação abrange o nível federal, com articulação da pauta esportiva com áreas como educação, assistência social, saúde e desenvolvimento territorial."),
    "Manual": ("Manual (saúde)",
               "A Manual (saúde) é uma plataforma digital voltada principalmente à saúde e bem-estar masculino, oferecendo atendimento online e tratamentos baseados em evidências (como saúde capilar, sono e saúde sexual), com prescrição médica e acompanhamento remoto. Possui atuação aprofundada em emagrecimento (foco em GLP-1 e redutores de apetite), disfunção erétil (consultas médicas e medicamentos manipulados) e queda capilar (Finasterida e Minoxidil). Tem interesse em promover inovação na área de saúde, principalmente em relação a manipuláveis e na expansão da telemedicina."),
    "Cactus": ("Instituto Cactus",
               "O Instituto Cactus é uma organização filantrópica e de direitos humanos, sem fins lucrativos e independente, que atua para ampliar e qualificar o ecossistema da saúde mental no Brasil, com foco prioritário em mulheres e adolescentes. Sua atuação organiza-se em duas frentes: o fomento estratégico (grant-making), financiando e co-criando iniciativas em saúde mental e produzindo evidências; e o advocacy, com foco na formulação, implementação e avaliação de políticas públicas e análise de projetos de lei. O Instituto também desenvolve ferramentas de apoio a gestores e promove ações de educação e mobilização social para reduzir o estigma em torno da saúde mental."),
    "Vital Strategies": ("Vital Strategies",
                         "A Vital Strategies é uma organização global de saúde pública que trabalha com governos e sociedade civil na concepção e implementação de políticas baseadas em evidências em áreas como doenças crônicas, segurança viária, qualidade do ar, dados vitais e comunicação de risco. Desde o ano passado tem focado na Reforma Tributária, em especial no Imposto Seletivo, buscando incidir sobre a alíquota em bebidas açucaradas, álcool e tabaco. Também trata do cuidado no trânsito relacionado ao uso de drogas, políticas de vedação de marketing de cigarros, dispositivos eletrônicos, ultraprocessados e bebidas alcoólicas. Na área tecnológica, investe em estudos que ligam Inteligência Artificial à jornada do paciente, com foco no combate ao feminicídio e no diagnóstico precoce de câncer. Acompanha ainda temas de saúde ambiental, qualidade do ar e intoxicação por chumbo."),
    "Mevo": ("Mevo",
             "A Mevo é uma healthtech brasileira que integra soluções de saúde digital, da prescrição eletrônica à compra e entrega de medicamentos, conectando médicos, hospitais, farmácias e pacientes para tornar o cuidado mais simples, eficiente e rastreável. Seu foco está na construção de um ecossistema digital interoperável, atuando junto aos Poderes Legislativo e Executivo para contribuir com o fortalecimento de uma Rede Nacional de Dados em Saúde (RNDS) robusta e integrada. Também mantém diálogo com agências reguladoras, acompanhando debates sobre saúde digital, interoperabilidade de sistemas, proteção de dados e normativas que impactem suas soluções tecnológicas."),
    "Coletivo Feminista": ("Coletivo Feminista",
                           "O Coletivo Feminista (Nem Presa Nem Morta) é um movimento feminista que atua pela descriminalização e legalização do aborto no Brasil, articulando pesquisa, incidência política e mobilização social. Seus princípios ético-políticos abrangem a comunicação como direito e fundamento da democracia, a defesa do Estado democrático de direito, a compreensão de que maternidade não é dever e deve respeitar a liberdade de escolha, a promoção de uma atenção universal, equânime e integral à saúde — com ênfase no papel do SUS, no acesso a métodos contraceptivos e abortivos seguros e no respeito à autodeterminação reprodutiva —, além da defesa da descriminalização e legalização do aborto. Desde o final de 2024, o coletivo tem focado no novo Código Civil e no PLD3/2025, que susta a resolução 258 do CONANDA, atuando para evitar regressões inconstitucionais ligadas ao aborto, em especial quando se trata de crianças e adolescentes."),
    "IDEC": ("Instituto Brasileiro de Defesa do Consumidor (Idec)",
             "O Instituto Brasileiro de Defesa do Consumidor (Idec) é uma associação civil sem fins lucrativos e independente de empresas, partidos ou governos, fundada em 1987. Atua na defesa dos direitos dos consumidores e na promoção de relações de consumo éticas, seguras e sustentáveis, combinando advocacy, pesquisa e litigância estratégica, com foco em temas como saúde, alimentação, energia, telecomunicações e direitos digitais. Destaca-se na incidência em políticas públicas relacionadas à alimentação saudável, controle de ultraprocessados e agrotóxicos, rotulagem nutricional, transição energética justa e regulação de plataformas digitais. Também acompanha a regulação dos planos de saúde junto à ANS, a saúde digital do ponto de vista do consumidor, greenwashing, práticas abusivas de telemarketing e debates como o ReData."),
    "Umane": ("Umane",
              "A Umane é uma organização da sociedade civil isenta, apartidária e sem fins lucrativos que atua para fomentar a saúde pública de forma sistêmica no Brasil, com foco em ampliar equidade, eficiência e qualidade do sistema de saúde. Opera por meio de fomento a projetos, articulação com parceiros e um modelo que combina monitoramento e avaliação, uso de dados e tecnologia (telessaúde e IA) e advocacy. Suas frentes programáticas incluem o fortalecimento da Atenção Primária à Saúde (APS), a atenção integral às Doenças Crônicas Não Transmissíveis (DCNT) — com foco em doenças cardiovasculares, diabetes tipo 2, obesidade, subnutrição e dislipidemias — e a saúde da mulher, da criança e do adolescente, com ênfase no pré-natal, nos primeiros mil dias e no enfrentamento da má nutrição infantil e juvenil."),
    "ASBAI": ("Associação Brasileira de Alergia e Imunologia (ASBAI)",
              "A Associação Brasileira de Alergia e Imunologia (ASBAI) é uma entidade científica sem fins lucrativos que reúne médicos especialistas em alergia e imunologia clínica no Brasil. Atua na promoção do ensino, pesquisa e atualização profissional, elaborando diretrizes clínicas, posicionamentos técnicos e recomendações para o diagnóstico e tratamento de doenças alérgicas e imunológicas. Acompanha debates regulatórios junto ao Ministério da Saúde (especialmente a Conitec) e à Anvisa, em temas como incorporação de tecnologias, imunobiológicos, vacinas e protocolos clínicos. Sua principal linha de atuação neste momento é a incorporação da caneta de adrenalina autoinjetável no SUS e a obrigatoriedade de notificação de ocorrências de anafilaxia ao Ministério da Saúde."),
    "Infinis": ("Instituto Futuro é Infância Saudável (Infinis)",
                "O Instituto Futuro é Infância Saudável (Infinis) é a frente de filantropia estratégica e advocacy da Fundação José Luiz Setúbal (FJLS). A organização atua com base em evidências científicas para promover políticas públicas, fortalecer a sociedade civil e impulsionar soluções que assegurem saúde e bem-estar na infância. Sua atuação está estruturada em quatro eixos temáticos: segurança alimentar e enfrentamento da má nutrição; saúde mental; prevenção às violências; e fortalecimento da sociedade civil, alinhados aos ODS da ONU. Com foco na incidência política, busca contribuir para o aprimoramento e a efetiva implementação de políticas públicas, além de fomentar a transformação de comportamentos e o desenvolvimento de soluções locais sustentáveis."),
}

PROMPT = Template(
"""Você é analista de políticas públicas e faz triagem de atos do DOU, matérias legislativas e notícias para um(a) cliente.

Missão/escopo do cliente:
$descricao

Tarefa:
Classificar o alinhamento do **Conteúdo** com a missão do cliente.

Regras de evidência:
- Use **apenas** o Conteúdo e a descrição do cliente acima. Não utilize conhecimento próprio sobre o cliente além do que está descrito neste prompt.
- NÃO exija que o Conteúdo cubra TODA a missão do cliente.
  # Se o Conteúdo estiver claramente dentro de ao menos UMA frente/eixo relevante do cliente, marque "Alinha".
  # A ausência de menção a outras frentes (ex.: socioemocional) NÃO reduz automaticamente para "Parcial".
- Use "Parcial" apenas nos dois casos descritos abaixo — não como classe-padrão para dúvidas genéricas.
- Se o texto for claramente de natureza incompatível com triagem temática (ex.: decisão sobre caso individual sem política pública; deferimento/indeferimento nominal; concessão pontual; nomeação/dispensa rotineira sem tema; mero expediente administrativo sem objeto; publicação que não permite inferir assunto), marque "Não se aplica".

Classes (escolha exatamente UMA):
- "Alinha": O objeto/tema do Conteúdo é claro e há evidência explícita de relação com pelo menos 1 frente/eixo do cliente.
- "Parcial": Use SOMENTE em um destes dois casos:
    (a) Ambiguidade temática — o Conteúdo trata de tema que poderia ou não se encaixar na missão, mas o texto é insuficiente para decidir com segurança;
    (b) Cobertura incompleta — o Conteúdo aborda parcialmente o tema do cliente, mas mistura substancialmente outras agendas não relacionadas, de modo que a relevância é real porém limitada.
- "Não Alinha": O tema é claro e não tem relação com a missão do cliente.
- "Não se aplica": O Conteúdo não é classificável por tema/escopo do cliente com base no texto (exemplos acima), ou é predominantemente ato individual/procedimental sem política pública inferível.

Restrições estruturais obrigatórias:
- NÃO inclua qualquer metacomentário sobre a classificação.
- NÃO mencione que está classificando, analisando ou respondendo ao prompt.
- A justificativa deve conter APENAS uma descrição objetiva do que o Conteúdo trata (objeto/tema).
- NÃO explique impactos potenciais, intenções do autor ou interpretações jurídicas.
- NÃO utilize linguagem avaliativa ou argumentativa.

Formato de saída:
Retorne **somente** JSON válido neste formato:
{
  "alinhamento": "Alinha" | "Parcial" | "Não Alinha" | "Não se aplica",
  "justificativa": "1–3 frases citando elementos do Conteúdo (termos/trechos) que sustentam a decisão"
}

Conteúdo:
<conteudo>
$conteudo
</conteudo>""".strip()
)

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
    # Sanitiza o conteúdo para não quebrar o delimitador XML do prompt.
    conteudo_safe = conteudo.replace("</conteudo>", "</conteudo\u200b>")
    prompt_text = PROMPT.substitute(descricao=descricao_cliente, conteudo=conteudo_safe)
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

    _, desc_cli = CLIENTE_DESCRICOES.get(title, (title, ""))

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
