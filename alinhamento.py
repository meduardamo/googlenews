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
    "IU": (
        "Instituto Unibanco (IU)",
        "O Instituto Unibanco (IU) é uma organização sem fins lucrativos que atua no fortalecimento da gestão educacional, desenvolvendo projetos como o Jovem de Futuro, oferecendo apoio técnico a secretarias estaduais de educação e produzindo conhecimento para aprimorar políticas públicas. Seu foco está tanto no cenário federal, acompanhando os debates sobre o financiamento da educação, programas nacionais de educação, regulação educacional e diretrizes definidas por órgãos como o Conselho Nacional de Educação, quanto subnacional, olhando para 6 estados prioritários (RS, MG, ES, CE, PI e GO). O IU apoia iniciativas de recomposição de aprendizagens, infraestrutura escolar, inclusão digital, educação ambiental, mudanças do clima e valorização de profissionais da educação.",
    ),
    "FMCSV": (
        "Fundação Maria Cecilia Souto Vidigal (FMCSV)",
        "A Fundação Maria Cecilia Souto Vidigal (FMCSV) é uma organização da sociedade civil dedicada ao fortalecimento da primeira infância no Brasil. Sua atuação concentra-se na integração entre produção de conhecimento, advocacy e apoio à formulação e implementação de políticas públicas, com o objetivo de assegurar o desenvolvimento integral de crianças de 0 a 6 anos. A Fundação acompanha o debate sobre educação domiciliar (homeschooling), posicionando-se de forma contrária a avanços nessa pauta. Além disso, participa ativamente da construção e implementação da Política Nacional Integrada da Primeira Infância. Desde 2007, a Fundação trabalha para garantir que todas as crianças brasileiras tenham uma infância saudável, com seus direitos plenamente assegurados. Com o lançamento da Agenda 2030 pela Organização das Nações Unidas (ONU), a instituição passou a alinhar suas estratégias à meta 4.2 dos Objetivos de Desenvolvimento Sustentável (ODS), que trata da garantia de acesso a cuidados e educação de qualidade na primeira infância. Entre suas iniciativas, destaca-se o programa 'Primeira Infância Primeiro', que disponibiliza dados, evidências e ferramentas para gestores públicos e candidatos, contribuindo para a qualificação do debate e das políticas voltadas à infância.",
    ),
    "IEPS": (
        "Instituto de Estudos para Políticas de Saúde (IEPS)",
        "O Instituto de Estudos para Políticas de Saúde (IEPS) é uma organização independente e sem fins lucrativos dedicada a aprimorar políticas de saúde no Brasil, combinando pesquisa aplicada, produção de evidências e advocacy em temas como atenção primária, saúde digital e financiamento do SUS. Com especialização em políticas públicas de saúde, o IEPS possui uma atuação centrada no fortalecimento do SUS enquanto sistema universal e equitativo. A organização se expande pela observação dos mais diversos temas, assim, é complexo delimitar seus núcleos de observação. No âmbito da análise da organização e governança federativa do SUS, com atenção especial ao modelo tripartite (União–Estados–Municípios), ao papel do Ministério da Saúde e às distorções introduzidas por emendas parlamentares no orçamento setorial. Existe um foco técnico na Estruturação da Atenção Primária à Saúde (APS), financiamento per capita e critérios redistributivos, regionalização como instrumento de coordenação federativa, e planejamento e alocação eficiente de recursos. Ademais, existe uma busca por equidade e enfrentamento de desigualdades, através do monitoramento de políticas criadas com ênfase em grupos historicamente vulnerabilizados, população negra, povos indígenas e originários, pessoas com deficiência, população LGBTQIA+, pessoas em situação de rua, e crianças, adolescentes, mulheres, homens e idosos. No campo da força de trabalho em saúde, a instituição aborda a relação entre disponibilidade de profissionais, organização federativa e financiamento do SUS. A análise envolve a escassez e a distribuição territorial de médicos, enfermeiros e demais categorias, considerando desigualdades regionais e capacidade instalada dos entes subnacionais. É necessário destaque para o Programa Agora Tem Especialistas, que vem sendo acompanhado desde sua instituição no ano passado. No âmbito da saúde mental, a atuação contempla a organização da Rede de Atenção Psicossocial e a consolidação do processo de desinstitucionalização. A fiscalização de comunidades terapêuticas é um tema de alta relevância. O uso terapêutico de cannabis e derivados é tratado no âmbito da regulação sanitária. Seguindo a nova ordem de prioridades de atuação do IEPS, especialmente em seu trabalho de incidência na Secretaria Executiva da Frente Parlamentar Mista para a Promoção da Saúde Mental, as BETs são destaque no monitoramento para 2026. A organização também realiza o acompanhamento das decisões da Anvisa e da ANS, com foco na regulação de produtos, serviços e operadoras de planos. No eixo de emergências sanitárias, a instituição acompanha políticas de vigilância epidemiológica, declaração de estados de emergência e coordenação federativa em situações de crise. Como as mudanças climáticas vêm sendo consideradas no planejamento das políticas de saúde, são observados debates relacionados a eventos extremos, deslocamentos populacionais e impactos sobre doenças transmissíveis e crônicas.",
    ),
    "IAS": (
        "Instituto Ayrton Senna (IAS)",
        "O Instituto Ayrton Senna (IAS) é um centro de inovação em educação que atua em pesquisa e desenvolvimento, disseminação em larga escala e influência em políticas públicas, com foco em aprendizagem acadêmica e competências socioemocionais na rede pública.",
    ),
    "ISG": (
        "Instituto Sonho Grande (ISG)",
        "O Instituto Sonho Grande (ISG) é uma organização sem fins lucrativos e apartidária voltada à expansão e qualificação do ensino médio integral em redes públicas. Atua em parceria com secretarias estaduais de educação, oferecendo apoio na revisão curricular, na formação de equipes escolares e na implementação de modelos de gestão orientados a resultados. Além disso, o Instituto mantém foco no fortalecimento da infraestrutura das escolas públicas de tempo integral e na promoção de políticas voltadas à alfabetização. Também acompanha o debate sobre educação domiciliar (homeschooling), posicionando-se de forma contrária à sua ampliação. Em colaboração com governos estaduais e organizações do terceiro setor, o ISG trabalha para ampliar o acesso ao Ensino Médio Integral em escolas públicas. Por meio de pesquisas e avaliações contínuas, analisa os impactos desse modelo educacional para a formação de jovens.",
    ),
    "Reúna": (
        "Instituto Reúna",
        "O Instituto Reúna desenvolve pesquisas e ferramentas para apoiar redes e escolas na implementação de políticas educacionais alinhadas à BNCC, com foco em currículo, materiais de apoio e formação de professores.",
    ),
    "REMS": (
        "Rede Esporte pela Mudança Social (REMS)",
        "A REMS – Rede Esporte pela Mudança Social articula organizações que usam o esporte como vetor de desenvolvimento humano, mobilizando atores e produzindo conhecimento para ampliar o impacto social dessa agenda no país. A rede atua desde 2007 no fortalecimento do campo do esporte para o desenvolvimento social, promovendo a troca de experiências, a sistematização de práticas e a realização de agendas coletivas, bem como acompanhando e incidindo em debates sobre políticas públicas, financiamento, marcos regulatórios e programas governamentais relacionados ao esporte educacional, comunitário e de participação. Sua atuação abrange o nível federal, com foco na qualificação da formulação e implementação de iniciativas, no fortalecimento técnico das organizações integrantes e no diálogo com gestores públicos e parlamentares, além da articulação da pauta esportiva com áreas como educação, assistência social, saúde e desenvolvimento territorial.",
    ),
    "Manual": (
        "Manual",
        "A Manual se posiciona como uma empresa de cuidado contínuo e personalizado, com foco em acesso facilitado, discrição e conveniência. Ela é uma plataforma digital voltada principalmente à saúde e bem-estar masculino, oferecendo atendimento online e tratamentos baseados em evidências (como saúde capilar, sono e saúde sexual), com prescrição médica e acompanhamento remoto. Tem um interesse em promover a inovação dentro da área de saúde, principalmente em relação a manipuláveis, como o princípio ativo GLP-1. Possui uma atuação aprofundada conectando clientes com médicos e tratamentos para emagrecimento (foco em GLP-1 e redutores de apetite), disfunção erétil (oferecendo serviços que incluem consultas médicas e medicamentos manipulados que impedem a ação da enzima PDE 5) e queda capilar (acompanhamento junto com o uso de Finasterida e Minoxidil). Por serem uma plataforma, também possuem interesse na expansão da telemedicina e inovações no campo tecnológico associado à saúde.",
    ),
    "Cactus": (
        "Instituto Cactus",
        "O Instituto Cactus é uma organização filantrópica e de direitos humanos, sem fins lucrativos e independente, que atua para ampliar e qualificar o ecossistema da saúde mental no Brasil, desenvolvendo projetos voltados à prevenção de agravos e à promoção do cuidado, com foco prioritário em mulheres e adolescentes. Sua atuação se organiza em duas frentes complementares: o fomento estratégico (grant-making), por meio do qual financia, co-cria e oferece suporte técnico a iniciativas que constroem e ampliam soluções e ferramentas em saúde mental, além de produzir evidências e incentivar inovações no campo da atenção psicossocial; e o advocacy, com foco na formulação, implementação e avaliação de políticas públicas, bem como na análise qualificada de projetos de lei. O Instituto também realiza incidência política para fortalecer a agenda da saúde mental no debate público e institucional, desenvolve ferramentas de apoio a gestores e governos e promove ações de educação, sensibilização e mobilização social, que objetivam reduzir o estigma e consolidar uma narrativa mais humanizada sobre o tema no país.",
    ),
    "Vital Strategies": (
        "Vital Strategies",
        "A Vital Strategies é uma organização global de saúde pública que trabalha com governos e sociedade civil na concepção e implementação de políticas baseadas em evidências em áreas como doenças crônicas, segurança viária, qualidade do ar, dados vitais e comunicação de risco. A organização trabalha com base em dados para a saúde e mortes evitáveis, então os temas são muito intersetoriais. Desde o ano passado tem focado na Reforma Tributária, em especial no Imposto Seletivo, buscando incidir sobre a alíquota em bebidas açucaradas, álcool e tabaco. O intuito é atingir um consumo zero sobre conteúdos que geram malefícios à saúde, como no caso de DCNTs, como hipertensão arterial. Para além desta campanha, também trata do cuidado no trânsito, observando acidentes que estão relacionados ao uso de drogas (lícitas ou não). Políticas sobre vedação total ou parcial de marketing, publicidade e rotulagem de cigarros, dispositivos eletrônicos para fumar, alimentos ultraprocessados e bebidas alcoólicas também são de interesse. Na área de desenvolvimento tecnológico, tem investido nos estudos que ligam Inteligência Artificial à jornada do paciente, buscando dois pontos principais: combate ao feminicídio e diagnóstico precoce de câncer. Com a incidência sobre a COP30 no ano passado, temas como saúde ambiental, qualidade do ar e intoxicação por chumbo também são acompanhados pela organização.",
    ),
    "Mevo": (
        "Mevo",
        "A Mevo é uma healthtech brasileira que integra soluções de saúde digital, da prescrição eletrônica à compra e entrega de medicamentos, conectando médicos, hospitais, farmácias e pacientes para tornar o cuidado mais simples, eficiente e rastreável. Seu foco está na construção de um ecossistema digital interoperável, atuando de forma contínua junto aos Poderes Legislativo e Executivo para contribuir com o fortalecimento de uma Rede Nacional de Dados em Saúde (RNDS) robusta e integrada. A empresa também mantém diálogo com agências reguladoras, com o objetivo de promover um ambiente normativo que viabilize uma rede de dados e de prescrição eletrônica interoperável e de alcance universal. Nesse contexto, acompanha e contribui para debates regulatórios relacionados à saúde digital, interoperabilidade de sistemas, proteção de dados e normativas que impactem o funcionamento de suas soluções tecnológicas.",
    ),
    "Coletivo Feminista": (
        "Coletivo Feminista",
        "O Nem Presa Nem Morta é um movimento feminista que atua pela descriminalização e legalização do aborto no Brasil, articulando pesquisa, incidência política e mobilização social. Seus princípios ético-políticos abrangem a comunicação como direito e fundamento da democracia, a defesa do Estado democrático de direito, a compreensão de que maternidade não é dever e deve respeitar a liberdade de escolha, a promoção de uma atenção universal, equânime e integral à saúde — com ênfase no papel do SUS, no acesso a métodos contraceptivos e abortivos seguros e no respeito à autodeterminação reprodutiva —, além da defesa da descriminalização e legalização do aborto. Desde o final do ano passado, o coletivo tem focado em dois projetos essenciais: o novo Código Civil e o PLD3/2025, que susta a resolução 258 do Conselho Nacional dos Direitos da Criança e do Adolescente (CONANDA). Em ambos os projetos a atuação da organização tem sido em evitar regressões inconstitucionais ligadas ao aborto, em especial quando se trata de crianças e adolescentes.",
    ),
    "IDEC": (
        "Instituto Brasileiro de Defesa do Consumidor (Idec)",
        "O Instituto Brasileiro de Defesa do Consumidor (Idec) é uma associação civil sem fins lucrativos, fundada em 1987, independente de empresas, partidos ou governos, que atua na defesa dos direitos dos consumidores e na promoção de relações de consumo éticas, seguras e sustentáveis. Sua atuação combina advocacy, pesquisa e litigância estratégica, com foco em temas como saúde, alimentação, energia, telecomunicações e direitos digitais, sendo os temas pautados muitas vezes transversais a todas as áreas. O Idec se destaca na formulação e incidência em políticas públicas relacionadas à promoção da alimentação adequada e saudável, ao controle de ultraprocessados e agrotóxicos, à rotulagem nutricional, à transição energética justa e à regulação de plataformas digitais. Também acompanha de perto a regulação dos planos de saúde, atuando junto à ANS, e a saúde digital do ponto de vista do direito do consumidor. Pauta temas como greenwashing e práticas abusivas de telemarketing. Paralelamente, produz estudos, materiais técnicos e eventos voltados à informação e mobilização da sociedade, mantendo diálogo com o Legislativo por meio de parcerias e incidência em projetos de lei, inclusive em debates como os relacionados ao ReData.",
    ),
    "Umane": (
        "Umane",
        "A Umane é uma organização da sociedade civil isenta, apartidária e sem fins lucrativos que atua para fomentar a saúde pública de forma sistêmica no Brasil, com foco em ampliar equidade, eficiência e qualidade do sistema de saúde. Sua missão é apoiar iniciativas transformadoras de prevenção de doenças e promoção da saúde que melhorem a qualidade de vida da população, operando por meio de fomento a projetos, articulação com uma rede de parceiros e um modelo de trabalho que combina monitoramento e avaliação, uso de dados e tecnologia (como telessaúde e uso de IA, com foco sempre na inovação dentro da área de saúde) e advocacy/comunicação para fortalecer políticas públicas. As frentes programáticas explicitadas pela Umane incluem o fortalecimento da Atenção Primária à Saúde (APS), a atenção integral às Doenças Crônicas Não Transmissíveis (DCNT), sendo o foco as doenças cardiovasculares, diabetes tipo 2, obesidade, subnutrição e dislipidemias; e a saúde da mulher, da criança e do adolescente, com ênfase na articulação entre os níveis de atenção à saúde para o pré-natal, no acompanhamento integral dos primeiros mil dias e no enfrentamento da má nutrição infantil e juvenil.",
    ),
    "ASBAI": (
        "Associação Brasileira de Alergia e Imunologia (ASBAI)",
        "A Associação Brasileira de Alergia e Imunologia (ASBAI) é uma entidade científica sem fins lucrativos que reúne médicos especialistas em alergia e imunologia clínica no Brasil. Atua na promoção do ensino, pesquisa e atualização profissional nessas áreas, elaborando diretrizes clínicas, posicionamentos técnicos e recomendações para o diagnóstico e tratamento de doenças alérgicas e imunológicas. Seu foco está tanto no cenário nacional, acompanhando debates regulatórios junto ao Ministério da Saúde (especialmente a Conitec) e à Anvisa, especialmente em temas como incorporação de tecnologias, imunobiológicos, vacinas, assistência farmacêutica e protocolos clínicos, quanto na articulação com sociedades médicas estaduais e internacionais. A ASBAI também promove congressos, cursos e campanhas de conscientização sobre condições como asma, rinite alérgica, dermatite atópica, alergias alimentares, imunodeficiências primárias e anafilaxia. Sua principal linha de atuação neste momento é sobre a incorporação da caneta de adrenalina autoinjetável no SUS e também a obrigatoriedade de notificação ao Ministério da Saúde de ocorrências de anafilaxia/choque anafilático.",
    ),
    "Infinis": (
        "Instituto Futuro é Infância Saudável (Infinis)",
        "O Instituto Futuro é Infância Saudável (Infinis) é a frente de filantropia estratégica e advocacy da Fundação José Luiz Setúbal (FJLS). A organização atua com base em evidências científicas para promover políticas públicas, fortalecer a sociedade civil e impulsionar soluções que assegurem saúde e bem-estar na infância. Sua atuação está estruturada em quatro eixos temáticos: segurança alimentar e enfrentamento da má nutrição; saúde mental; prevenção às violências; e fortalecimento da sociedade civil. Esses eixos estão alinhados aos Objetivos de Desenvolvimento Sustentável (ODS) da ONU, especialmente no que se refere à promoção da saúde, da equidade e da proteção de crianças e adolescentes. Com foco na incidência política, o Infinis busca contribuir para o aprimoramento e a efetiva implementação de políticas públicas, além de fomentar a transformação de comportamentos e o desenvolvimento de soluções locais sustentáveis. No campo do fortalecimento da sociedade civil, apoia a produção de pesquisas científicas e o desenvolvimento de organizações de infraestrutura que atuam no setor.",
    ),
    "UNFPA": (
        "Fundo de População das Nações Unidas (UNFPA)",
        "O Fundo de População das Nações Unidas (UNFPA) é a agência das Nações Unidas voltada à saúde sexual e reprodutiva e às questões de população e desenvolvimento. No Brasil, atua desde 1973 em cooperação com governos, organismos internacionais, sociedade civil e outros parceiros para apoiar a formulação, implementação e monitoramento de políticas públicas baseadas em direitos, evidências e redução de desigualdades. Sua missão é contribuir para um mundo em que todas as gestações sejam desejadas, todos os partos sejam seguros e cada jovem alcance seu potencial. A organização acompanha temas como saúde sexual e reprodutiva, direitos reprodutivos, planejamento reprodutivo, mortalidade materna, atenção obstétrica, acesso a contraceptivos e insumos de saúde, gravidez na adolescência, uniões infantis, violência baseada em gênero, feminicídio, exploração sexual, tráfico de pessoas, juventudes, juventude negra, igualdade racial, população e produção de dados para políticas públicas. Também atua em agendas transversais, como emergências humanitárias, mudanças climáticas, justiça climática, gênero e proteção de populações vulnerabilizadas. No campo da incidência pública, o UNFPA tende a priorizar debates relacionados à autonomia corporal, à garantia de direitos de meninas, mulheres, adolescentes e jovens, ao enfrentamento de desigualdades raciais e territoriais e ao uso de dados demográficos, censitários e populacionais para subsidiar a formulação, implementação e monitoramento de políticas públicas.",
    ),
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
