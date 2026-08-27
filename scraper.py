import os, re, time, sys, random
from datetime import datetime, timedelta, date

import feedparser
import pandas as pd
from newspaper import Article

import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials

# Limites e flags
MAX_CELL_CHARS = int(os.getenv("MAX_CELL_CHARS", "47000"))
# Teto de linhas por aba. O que derruba a planilha ("This document is too large
# to continue editing") não é a contagem de células — medida em 27/07/2026 ela
# estava em 1,19 milhão, 12% do limite de 10 milhões — e sim o VOLUME DE TEXTO:
# ~716 milhões de caracteres. Com o resumo limpo (abaixo) cada linha cai de
# ~12.000 para ~2.000 caracteres, e o teto de linhas segura o crescimento.
LIMITE_LINHAS_ABA = int(os.getenv("LIMITE_LINHAS_ABA", "12000"))

# Teto do resumo. O feed RSS entrega em 'summary' o artigo inteiro em HTML
# (inclusive <img>), com 9.400 caracteres em média: 78% do peso do documento.
# Resumo é para ler de relance; o corpo tem coluna própria.
MAX_RESUMO_CHARS = int(os.getenv("MAX_RESUMO_CHARS", "1200"))

# Teto do corpo. O alinhamento classifica se a notícia toca a missão do cliente;
# alguns milhares de caracteres bastam para isso. Sem teto, a reserva do RSS
# abaixo só transferiria o peso de 'resumo' para 'texto_completo'.
MAX_TEXTO_CHARS = int(os.getenv("MAX_TEXTO_CHARS", "4000"))
SHEETS_SPLIT_TEXT = os.getenv("SHEETS_SPLIT_TEXT", "0").strip() in ("1","true","True","yes","on")

# ── retry helper ──────────────────────────────────────────────────────────────
def _gsheets_with_retry(fn, *args, max_retries=6, **kwargs):
    """Chama fn(*args, **kwargs) com backoff exponencial em caso de 429."""
    for attempt in range(max_retries):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = getattr(e.response, 'status_code', None)
            if status == 429 and attempt < max_retries - 1:
                wait = (2 ** attempt) + random.uniform(0, 1)
                print(f"  ⚠ Rate limited (429). Aguardando {wait:.1f}s (tentativa {attempt+1}/{max_retries})...")
                time.sleep(wait)
            else:
                raise

# ── helpers de limpeza/limite ─────────────────────────────────────────────────
def _clean_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    return "".join(ch for ch in s if (ch == "\n") or ((ord(ch) >= 32) and (ord(ch) != 127)))

_TAG_HTML = re.compile(r"<[^>]+>")
_ESPACOS = re.compile(r"[ \t\r\f\v]+")


def _sem_html(s: str) -> str:
    """Tira marcação e entidades do 'summary' do RSS.

    Desescapa ANTES de remover as tags, e repete: parte dos feeds entrega o HTML
    escapado (&lt;p&gt;), e limpar na ordem inversa deixaria a tag aparecer no
    resultado final, já sem chance de ser removida.
    """
    import html as _html_mod
    texto = str(s or "")
    for _ in range(3):                       # &amp;lt;p&amp;gt; precisa de 2 voltas
        antes = texto
        texto = _html_mod.unescape(texto)
        texto = _TAG_HTML.sub(" ", texto)
        if texto == antes:
            break
    return _ESPACOS.sub(" ", texto).strip()


def _truncate(s: str, limit: int = MAX_CELL_CHARS) -> str:
    s = _clean_text(s)
    if len(s) <= limit:
        return s
    return s[:max(0, limit - 10)] + " [..]"

def _split_in_chunks(s: str, limit: int = MAX_CELL_CHARS):
    s = _clean_text(s)
    if not s:
        return [""]
    return [s[i:i + limit] for i in range(0, len(s), limit)]

def _enforce_sheet_limits(df: pd.DataFrame, limit: int = MAX_CELL_CHARS):
    for c in df.columns:
        if df[c].dtype == object or str(df[c].dtype).startswith("string"):
            df[c] = df[c].astype(str).map(lambda v: _truncate(v, limit))
    return df

# ── dados clientes ────────────────────────────────────────────────────────────
CLIENT_THEME_DATA = """
ISG|Educação|tempo integral; fundeb; ensino técnico profissionalizante; educação profissional e tecnológica; FNDE; ensino médio; propag; infraestrutura escolar; ensino fundamental integral; alfabetização integral; escola em tempo integral; programa escola em tempo integral; ensino fundamental em tempo integral; plano nacional de educação; programa pé-de-meia; PNEERQ; INEP; conselho nacional de educação; PDDE; SNE; celular nas escolas; conselhos de educação; programa de fomento às escolas de ensino médio em tempo integral
IU|Educação|recomposição da aprendizagem; educação em tempo integral; fundeb; educação e equidade; educação profissional e tecnológica; ensino técnico profissionalizante; FNDE; gestão educacional; diretores escolares; magistério; ensino médio; adaptações de escolas; educação ambiental; plano nacional de educação; PDDE; programa pé de meia; INEP; conselho nacional de educação; VAAT; VAAR; secretaria estadual de educação; celular nas escolas; EAD
REMS|Esportes|esporte; esporte e desenvolvimento social; esporte e educação; esporte e equidade; paradesporto; desenvolvimento social; esporte educacional; plano nacional de esporte; lei de incentivo ao esporte; sistema nacional de esporte; conselho nacional de esporte; esporte e saúde; esporte para toda a vida; esporte amador
FMCSV|Primeira infância|criança; criança feliz; alfabetização; creche; conanda; maternidade; parentalidade; paternidade; primeira infância; infantil; infância; infanto-juvenil; fundeb; educação básica; plano nacional de educação; homeschooling; FNDE; vaar; vaat; educação infantil; pré-escola; materno-infantil; infraestrutura escolar; política nacional de cuidados; bolsa família; visitação domiciliar
IEPS|Saúde|sus; sistema único de saúde; equidade em saúde; atenção primária à saúde; vigilância epidemiológica; planos de saúde; caps; seguros de saúde; populações vulneráveis; desigualdades sociais; organização do sus; políticas públicas em saúde; governança do sus; regionalização em saúde; população negra em saúde; saúde indígena; povos originários; saúde da pessoa idosa; envelhecimento ativo; atenção primária; saúde da criança; saúde do adolescente; saúde da mulher; saúde do homem; saúde da pessoa com deficiência; saúde da população lgbtqia+; financiamento da saúde; emendas e orçamento da saúde; emendas parlamentares; ministério da saúde; trabalhadores e profissionais de saúde; força de trabalho em saúde; política de recursos humanos em saúde; formação profissional de saúde; cuidados primários em saúde; emergências climáticas e ambientais em saúde; emergências climáticas; mudanças ambientais; adaptação climática; saúde ambiental; políticas climáticas; vigilância em saúde; epidemiológica; emergência em saúde; estado de emergência; saúde suplementar; seguradoras; planos populares; anvisa; ans; sandbox regulatório; cartões e administradoras de benefícios em saúde; economia solidária em saúde mental; pessoa em situação de rua; saúde mental; fiscalização de comunidades terapêuticas; rede de atenção psicossocial; raps; unidades de acolhimento; assistência multiprofissional; centros de convivência; cannabis; canabidiol; tratamento terapêutico; desinstitucionalização; manicômios; hospitais de custódia; saúde mental na infância; adolescência; escolas; comunidades escolares; protagonismo juvenil; dependência química; vícios; ludopatia; capacitação em saúde mental; intervenções terapêuticas em saúde mental; internet e redes sociais na saúde mental; violência psicológica; surto psicótico
Umane|Saúde|sus; sistema único de saúde; atenção primária à saúde; vigilância epidemiológica; planos de saúde; caps; equidade em saúde; populações vulneráveis; desigualdades sociais; organização do sus; políticas públicas em saúde; governança do sus; regionalização em saúde; população negra em saúde; saúde indígena; povos originários; saúde da pessoa idosa; envelhecimento ativo; atenção primária; saúde da criança; saúde do adolescente; saúde da mulher; saúde do homem; saúde da pessoa com deficiência; saúde da população lgbtqia+; financiamento da saúde; emendas e orçamento da saúde; emendas parlamentares; ministério da saúde; trabalhadores e profissionais de saúde; força de trabalho em saúde; política de recursos humanos em saúde; formação profissional de saúde; cuidados primários em saúde; emergências climáticas e ambientais em saúde; emergências climáticas; mudanças ambientais; adaptação climática; saúde ambiental; políticas climáticas; vigilância em saúde; epidemiológica; emergência em saúde; estado de emergência; saúde suplementar; seguradoras; planos populares; anvisa; ans; sandbox regulatório; cartões e administradoras de benefícios em saúde; conass; conasems
Cactus|Saúde|saúde mental; saúde mental para meninas; saúde mental para juventude; saúde mental para mulheres; pse; eca; rede de atenção psicossocial; raps; caps; centro de apoio psicossocial; programa saúde na escola; bullying; cyberbullying; eca digital
Vital Strategies|Saúde|saúde mental; dados para a saúde; morte evitável; doenças crônicas não transmissíveis; rotulagem de bebidas alcoólicas; educação em saúde; bebidas alcoólicas; imposto seletivo; rotulagem de alimentos; alimentos ultraprocessados; publicidade infantil; publicidade de alimentos ultraprocessados; tributação de bebidas alcoólicas; alíquota de bebidas alcoólicas; cigarro eletrônico; controle de tabaco; violência doméstica; exposição a fatores de risco; departamento de saúde mental; hipertensão arterial; saúde digital; violência contra crianças; violência contra mulheres; feminicídio; cop 30
ASBAI|Saúde|alergia e imunologia; ASBAI; doenças imunoalérgicas; alergistas; imunologistas; atenção especializada; atenção primária; estratégia saúde da família; ESF; SUS; residência médica; bolsas de residência; formação médica; titulação médica; CNRM; comissão mista de especialidades; SGTES; SAES; SAPS; PCDT; SVSA; anvisa; conitec; tabela SUS; linhas de cuidado; referência e contrarreferência; anafilaxia; choque anafilático; adrenalina autoinjetável; epinefrina autoinjetável; caneta de adrenalina; EpiPen; PL 85/2024; PL 793/2024; asma; rinite; alergia alimentar; dermatite atópica; urticária crônica; imunodeficiências; erros inatos da imunidade; doenças raras; vigilância epidemiológica; notificação compulsória; incorporação tecnológica; assistência farmacêutica; registro sanitário; regulação sanitária; inovação em saúde; ciência, tecnologia e inovação em saúde; complexo econômico-industrial da saúde; acesso a tecnologias em saúde
Infinis|Primeira infância|primeira infância; conanda; criança; estatuto da criança e do adolescente; Infinis; saúde mental; RAPS; violência contra crianças; violência psicológica; saúde indígena; plano nacional de enfrentamento da violência contra crianças e adolescentes; protagonismo juvenil; coalizão brasileira pelo fim da violência contra crianças e adolescentes; violência sexual contra crianças; sistema único de assistência social; pesquisa nacional de saúde mental; adolescente; insegurança alimentar; orçamento primeira infância; assistência social; saúde mental indígena
UNFPA|Saúde|casamento de menores; casamento de menores de 16 anos; idade mínima para casamento; nulidade do casamento; anulabilidade do casamento; união estável de menores; união estável adolescente; capacidade civil para casamento; prevenção ao casamento infantil; erradicação do casamento infantil; diretrizes para prevenção do casamento infantil; proteção de meninas adolescentes; enfrentamento das uniões informais precoces; prevenção de uniões envolvendo crianças e adolescentes; casamento infantil; casamento de adolescentes
Coletivo Feminista|Direitos reprodutivos|direitos reprodutivos; aborto; aborto legal; nascituro; criança; gestação acima de 22 semanas; interrupção legal da gestação; interrupção da gestação; resolução 258 conanda; vida por nascer; vida desde a concepção; desde a concepção; criança por nascer; infanticídio; feticídio; assistolia fetal; medicamento abortivo; misoprostol; citotec; cytotec; mifepristona; ventre; assassinato de bebês; luto parental; síndrome pós aborto; síndrome pós-aborto; gravidez; direitos das mulheres; violência sexual contra mulheres; violência sexual contra meninas; autonomia reprodutiva; planejamento reprodutivo; contracepção; objeção de consciência; escusa de consciência; natimorto; gravidez decorrente de estupro; gravidez infantil; gravidez na infância; maternidade forçada; gestação decorrente de estupro; menina mãe; menina-mãe
IDEC|Saúde|defesa do consumidor; ação civil pública; sac; reforma tributária; ultraprocessados; doenças crônicas não transmissíveis; dcnts; obesidade; codex alimentarius; gordura trans; adoçantes; edulcorantes; rotulagem de alimentos; transgênicos; organismos geneticamente modificados; ogms; marketing e publicidade de alimentos; comunicação mercadológica; escolas e alimentação escolar; bebidas açucaradas; refrigerante; programa nacional de alimentação escolar; pnae; educação alimentar e nutricional; agrotóxicos; pesticidas; defensivos fitossanitários; tributação de alimentos não saudáveis; desertos alimentares; desperdício de alimentos; segurança alimentar e nutricional; direito humano à alimentação; fome; sustentabilidade; mudança climática; plástico; gestão de resíduos; economia circular; desmatamento; greenwashing; energia elétrica; encargos tarifários; subsídios na tarifa de energia; descontos na tarifa de energia; energia pré-paga; abertura do mercado de energia para consumidor cativo; mercado livre de energia; qualidade do serviço de energia; serviço de energia; tarifa social de energia elétrica; geração térmica; combustíveis fósseis; transição energética; descarbonização da matriz elétrica; gases de efeito estufa; acordo de paris; objetivos do desenvolvimento sustentável; reestruturação do setor de energia; reforma do setor elétrico; modernização do setor elétrico; universalização do acesso à energia; eficiência energética; geração distribuída; carvão mineral; painel solar; crédito imobiliário; crédito consignado; publicidade de crédito; cartão de crédito; pagamento de fatura; parcelamento com e sem juros; cartões pré-pagos; programas de fidelidade; cheque especial; taxa de juros; contrato de crédito; endividamento de jovens; crédito estudantil; endividamento de idosos; crédito por meio de aplicativos; abertura e movimentação de conta bancária; cobrança de serviços sem autorização; cadastro positivo; contratação de serviços bancários com imposição de seguros e títulos de capitalização; acessibilidade aos canais de serviços bancários; serviços bancários; caixa eletrônico; internet banking; aplicativos móveis; contratação de pacotes de contas bancárias; acesso à informação em caso de negativa de crédito; plano de saúde; saúde suplementar; medicamentos isentos de prescrição; mip; medicamentos antibióticos; antimicrobianos; propriedade intelectual; patentes; licença compulsória; preços de medicamentos; complexo econômico-industrial da saúde; saúde digital; prontuário eletrônico; rede nacional de dados em saúde; rnds; datasus; proteção de dados pessoais; telessaúde; telecomunicações; internet; tv por assinatura; serviço de acesso condicionado; telefonia móvel; telefonia fixa; tv digital; lei geral de proteção de dados; autoridade nacional de proteção de dados; reconhecimento facial; lei geral de telecomunicações; bens reversíveis; fundo de universalização dos serviços de telecomunicações; provedores de acesso; franquia de dados; franquia de internet; marco civil da internet; neutralidade de rede; zero rating; privacidade; lei de acesso à informação; regulação de plataformas digitais; desinformação; fake news; dados biométricos; vazamento de dados; telemarketing; serviço de valor adicionado
MegaEdu|Educação|Conselho Gestor do FUST
""".strip()



def parse_client_theme_data(raw: str):
    mapping = {}
    for line in raw.splitlines():
        line = line.strip()
        if not line or '|' not in line:
            continue
        cliente, tema, keywords_str = [x.strip() for x in line.split('|', 2)]
        keywords = [k.strip() for k in keywords_str.split(';') if k.strip()]
        mapping[cliente] = {'tema': tema, 'keywords': keywords}
    return mapping


def make_wholeword_pattern(term: str):
    """Padrão de palavra inteira. Sigla em CAIXA ALTA casa com sensibilidade a
    maiúscula.

    Com IGNORECASE em tudo, "SUAS" (o sistema de assistência social) casava com
    o pronome "suas" e entrava em quase toda notícia: 81% do que ele trouxe foi
    classificado como Não Alinha. "CAPS" pegava "caps lock" pelo mesmo motivo.
    Sigla curta só vale quando escrita como sigla.
    """
    termo = term.strip()
    e_sigla = (len(termo) <= 6 and termo.isupper()
               and termo.isalpha() and len(termo) >= 2)
    flags = 0 if e_sigla else re.IGNORECASE
    return re.compile(rf'\b{re.escape(termo)}\b', flags=flags)


def parse_palavras_por_cliente_cell(cell: str):
    out = {}
    if not isinstance(cell, str) or not cell.strip():
        return out
    parts = [p.strip() for p in cell.split(';') if p.strip()]
    for part in parts:
        if '=' not in part:
            continue
        k, v = part.split('=', 1)
        k = k.strip()
        kws = [x.strip() for x in v.split('|') if x.strip()]
        out[k] = kws
    return out


class ColetorNoticias:
    def __init__(self, timezone_offset_hours=0, deduplicar=True):
        self.tz_offset = timedelta(hours=timezone_offset_hours)
        self.deduplicar = deduplicar
        self._seen_urls = set()
        self.df_noticias = pd.DataFrame(columns=[
            'data_publicacao', 'titulo', 'fonte', 'url',
            'palavras_por_cliente', 'resumo', 'texto_completo', 'data_coleta'
        ])
        self.clientes = parse_client_theme_data(CLIENT_THEME_DATA)
        self.cliente_patterns = {
            c: [make_wholeword_pattern(k) for k in data['keywords']]
            for c, data in self.clientes.items()
        }

        # Cache de worksheets: { spreadsheet_id: { ws_title: ws_object } }
        self._ws_cache = {}

        self.feeds = {
            # Base
            'G1 Política': 'https://g1.globo.com/rss/g1/politica/',
            'Folha Poder': 'https://feeds.folha.uol.com.br/poder/rss091.xml',
            'Estadão Política': 'https://politica.estadao.com.br/rss.xml',
            'O Globo Brasil': 'https://oglobo.globo.com/brasil/rss.xml',
            'UOL Notícias': 'https://rss.uol.com.br/feed/noticias.xml',
            'CNN Brasil': 'https://www.cnnbrasil.com.br/feed/',
            'Congresso em Foco': 'https://congressoemfoco.uol.com.br/feed/',
            'Poder360': 'https://www.poder360.com.br/feed/',
            'Metrópoles': 'https://www.metropoles.com/feed',
            'O Antagonista': 'https://www.oantagonista.com/feed/',
            'CartaCapital': 'https://www.cartacapital.com.br/feed/',
            'Nexo Jornal': 'https://www.nexojornal.com.br/rss.xml',
            'InfoMoney': 'https://www.infomoney.com.br/feed/',
            'Exame (via Google News)': 'https://news.google.com/rss/search?q=site:exame.com',
            'Money Times': 'https://www.moneytimes.com.br/feed/',
            'Agência Câmara': 'https://www.camara.leg.br/noticias/rss',
            'FolhaPE (home)': 'https://www.folhape.com.br/?format=feed&type=rss',
            'FolhaPE Política': 'https://www.folhape.com.br/politica/?format=feed&type=rss',
            'Diario de Pernambuco (via Google News)': 'https://news.google.com/rss/search?q=site:diariodepernambuco.com.br',
            'JC Online / NE10 (via Google News)': 'https://news.google.com/rss/search?q=site:jc.ne10.uol.com.br',

            # Política & Economia
            'Valor Econômico': 'https://valor.globo.com/rss/',
            'BBC Brasil – Política (topics)': 'https://www.bbc.com/portuguese/topics/cq23pdgvg85t/rss.xml',
            'Correio Braziliense': 'https://www.correiobraziliense.com.br/rss/',
            'Gazeta do Povo Política': 'https://www.gazetadopovo.com.br/politica/rss/',
            'Revista Veja Política': 'https://veja.abril.com.br/rss/politica.xml',
            'IstoÉ Política': 'https://istoe.com.br/assuntos/politica/feed/',

            # Governo/Institucional (via Google News)
            'Planalto (via GN)': 'https://news.google.com/rss/search?q=site:gov.br/planalto',
            'Casa Civil (via GN)': 'https://news.google.com/rss/search?q=site:gov.br/casacivil',
            'Ministério da Educação (via GN)': 'https://news.google.com/rss/search?q=site:gov.br/mec',
            'FNDE (via GN)': 'https://news.google.com/rss/search?q=site:gov.br/fnde',
            'Ministério da Saúde (via GN)': 'https://news.google.com/rss/search?q=site:gov.br/saude',
            'Anvisa (via GN)': 'https://news.google.com/rss/search?q=site:gov.br/anvisa',
            'Conass (via GN)': 'https://news.google.com/rss/search?q=site:conass.org.br',
            'Conasems (via GN)': 'https://news.google.com/rss/search?q=site:conasems.org.br',
            'INEP (via GN)': 'https://news.google.com/rss/search?q=site:inep.gov.br',
            'CNE (via GN)': 'https://news.google.com/rss/search?q=site:portal.mec.gov.br/cne',

            # Educação / Social
            'Revista Piauí': 'https://piaui.folha.uol.com.br/feed/',
            'Todos Pela Educação (via GN)': 'https://news.google.com/rss/search?q=site:todospelaeducacao.org.br',
            'Observatório do PNE (via GN)': 'https://news.google.com/rss/search?q=site:observatoriodopne.org.br',
            'CENPEC (via GN)': 'https://news.google.com/rss/search?q=site:cenpec.org.br',
            'Andifes (via GN)': 'https://news.google.com/rss/search?q=site:andifes.org.br',
            'CAPES (via GN)': 'https://news.google.com/rss/search?q=site:capes.gov.br',

            # Saúde
            'Fiocruz': 'https://portal.fiocruz.br/rss.xml',
            'Instituto Butantan (via GN)': 'https://news.google.com/rss/search?q=site:butantan.gov.br',
            'OPAS/OMS Brasil (via GN)': 'https://news.google.com/rss/search?q=site:paho.org/brasil',

            # Regionais
            'O Povo': 'https://www.opovo.com.br/rss/rss.xml',
            'Correio da Bahia': 'https://www.correio24horas.com.br/rss/',
            'Folha de Pernambuco': 'https://www.folhape.com.br/?format=feed&type=rss',

            # CNM
            'CNM (via Google News)': 'https://news.google.com/rss/search?q=site:cnm.org.br&hl=pt-BR&gl=BR&ceid=BR:pt-419',

            # CONSED
            'CONSED (via Google News)': 'https://news.google.com/rss/search?q=site:consed.org.br&hl=pt-BR&gl=BR&ceid=BR:pt-419',
        }

    # ── datas ──────────────────────────────────────────────────────────────────
    def _entry_datetime(self, entrada):
        dt = None
        if hasattr(entrada, 'published_parsed') and entrada.published_parsed:
            dt = datetime(*entrada.published_parsed[:6])
        elif hasattr(entrada, 'updated_parsed') and entrada.updated_parsed:
            dt = datetime(*entrada.updated_parsed[:6])
        elif 'published' in entrada:
            dt = pd.to_datetime(entrada.published, errors='coerce')
            dt = None if pd.isna(dt) else dt.to_pydatetime()
        elif 'updated' in entrada:
            dt = pd.to_datetime(entrada.updated, errors='coerce')
            dt = None if pd.isna(dt) else dt.to_pydatetime()
        return dt + self.tz_offset if dt else None

    def _eh_hoje(self, dt_obj):
        return bool(dt_obj) and dt_obj.date() == date.today()

    # ── matching ───────────────────────────────────────────────────────────────
    def _hits_por_cliente(self, texto):
        hits = {}
        base = texto or ""
        for cliente, patterns in self.cliente_patterns.items():
            ks = []
            for pat, kw in zip(patterns, self.clientes[cliente]['keywords']):
                if pat.search(base):
                    ks.append(kw)
            if ks:
                hits[cliente] = ks
        return hits

    def _format_palavras_por_cliente(self, hits_dict):
        return "; ".join(f"{cli}=" + "|".join(kws) for cli, kws in hits_dict.items())

    # ── dedup interno ──────────────────────────────────────────────────────────
    def noticia_existe(self, url):
        if not self.deduplicar:
            return False
        return (url in self._seen_urls) or (not self.df_noticias.empty and url in self.df_noticias['url'].values)

    def adicionar_noticia(self, noticia):
        if self.noticia_existe(noticia['url']):
            return False
        self._seen_urls.add(noticia['url'])
        self.df_noticias = pd.concat([self.df_noticias, pd.DataFrame([noticia])], ignore_index=True)
        return True

    # ── extração de corpo ──────────────────────────────────────────────────────
    def extrair_texto_completo(self, url):
        try:
            art = Article(url, language='pt')
            art.download()
            art.parse()
            return art.text
        except Exception:
            return ""

    # ── coleta ─────────────────────────────────────────────────────────────────
    def coletar_feeds(self, extrair_texto=False, apenas_hoje=True, pausa_seg=0.2):
        total_novas = 0
        print(f"\nIniciando coleta de {len(self.feeds)} fontes...\n" + "=" * 80)
        for nome_fonte, feed_url in self.feeds.items():
            print(f"\nProcessando: {nome_fonte}")
            try:
                feed = feedparser.parse(feed_url)
                novos = 0
                for e in feed.entries:
                    titulo = e.get('title', 'Sem título')
                    url = e.get('link', '')
                    resumo = _sem_html(e.get('summary', ''))[:MAX_RESUMO_CHARS]

                    dt_pub = self._entry_datetime(e)
                    if apenas_hoje and not self._eh_hoje(dt_pub):
                        continue

                    if extrair_texto:
                        corpo = self.extrair_texto_completo(url)
                        # newspaper3k volta quase vazio em site com paywall ou
                        # conteúdo por JS (o caso mais comum aqui: a coluna
                        # vinha com ~300 caracteres). O 'summary' do RSS, já
                        # limpo de HTML, é melhor corpo do que nada.
                        if len(corpo.strip()) < 400:
                            bruto = _sem_html(e.get('summary', ''))
                            if len(bruto) > len(corpo.strip()):
                                corpo = bruto
                        corpo = _sem_html(corpo)[:MAX_TEXTO_CHARS]
                        base_match = f"{titulo} {resumo} {corpo}"
                    else:
                        corpo = ''
                        base_match = f"{titulo} {resumo}"

                    hits_clientes = self._hits_por_cliente(base_match)
                    if not hits_clientes:
                        continue

                    noticia = {
                        'data_publicacao': dt_pub.strftime('%Y-%m-%d %H:%M:%S') if dt_pub else '',
                        'titulo': titulo,
                        'fonte': nome_fonte,
                        'url': url,
                        'palavras_por_cliente': self._format_palavras_por_cliente(hits_clientes),
                        'resumo': resumo,
                        'texto_completo': corpo,
                        'data_coleta': (datetime.now() + self.tz_offset).strftime('%Y-%m-%d %H:%M:%S'),
                    }

                    if self.adicionar_noticia(noticia):
                        novos += 1
                        total_novas += 1
                        if pausa_seg:
                            time.sleep(pausa_seg)
                print(f"  ✓ {novos} novas")
            except Exception as ex:
                print(f"  ✗ Erro em {nome_fonte}: {ex}")

        print("\n" + "=" * 80)
        print(f"Total novas: {total_novas} | Total na sessão: {len(self.df_noticias)}")
        return total_novas

    # ── Google Sheets ──────────────────────────────────────────────────────────
    def _gsheets_client(self):
        creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        return gspread.authorize(creds)

    def _build_ws_cache(self, sh):
        """Carrega todas as abas de uma vez (1 read request) e popula o cache."""
        worksheets = _gsheets_with_retry(sh.worksheets)
        self._ws_cache[sh.id] = {ws.title: ws for ws in worksheets}

    def _get_or_create_ws(self, sh, title):
        """Retorna worksheet pelo título, criando se necessário — sem read extra."""
        name = title[:31]

        # Garante que o cache do spreadsheet está populado
        if sh.id not in self._ws_cache:
            self._build_ws_cache(sh)

        if name not in self._ws_cache[sh.id]:
            ws = _gsheets_with_retry(sh.add_worksheet, title=name, rows=100, cols=20)
            self._ws_cache[sh.id][name] = ws

        return self._ws_cache[sh.id][name]

    def _somente_kws_do_cliente(self, cell, cliente):
        try:
            d = parse_palavras_por_cliente_cell(cell)
            return "|".join(d.get(cliente, []))
        except Exception:
            return ""

    def _read_ws_df(self, ws):
        values = _gsheets_with_retry(ws.get_all_values)
        if not values:
            return pd.DataFrame()
        header, rows = values[0], values[1:]
        if not any(h.strip() for h in header):
            return pd.DataFrame()
        width = len(header)
        norm_rows = [r + [""] * (width - len(r)) for r in rows]
        df = pd.DataFrame(norm_rows, columns=[h.strip() for h in header])
        return df

    def _ensure_header(self, ws, columns, min_rows=100, min_cols=None):
        values = _gsheets_with_retry(ws.get_all_values)
        has_header = bool(values and values[0] and any(v.strip() for v in values[0]))

        if min_cols is None:
            min_cols = max(20, len(columns))

        if not has_header:
            _gsheets_with_retry(ws.update, 'A1', [columns])

        if ws.row_count < max(2, min_rows):
            _gsheets_with_retry(ws.resize, rows=max(2, min_rows))
        if ws.col_count < min_cols:
            _gsheets_with_retry(ws.resize, cols=min_cols)

        return not has_header

    def exportar_por_cliente_para_sheets(self, spreadsheet_id: str, apenas_hoje=True):
        if self.df_noticias.empty:
            print("Sem notícias para exportar.")
            return

        df = self.df_noticias.copy()
        if apenas_hoje:
            hoje = date.today().strftime('%Y-%m-%d')
            df = df[df['data_publicacao'].astype(str).str.startswith(hoje)]

        gc = self._gsheets_client()
        sh = gc.open_by_key(spreadsheet_id)

        # Carrega todas as abas de uma vez → apenas 1 read request
        self._build_ws_cache(sh)

        # Sanitiza campos base
        for col in ['titulo', 'resumo', 'fonte', 'url']:
            if col in df.columns:
                df[col] = df[col].astype(str).map(_truncate)

        if 'texto_completo' in df.columns:
            if SHEETS_SPLIT_TEXT:
                df['texto_completo'] = df['texto_completo'].astype(str).map(_clean_text)
            else:
                df['texto_completo'] = df['texto_completo'].astype(str).map(_truncate)

        falhas_abas = []
        for cliente in self.clientes.keys():
            mask = df['palavras_por_cliente'].astype(str).str.contains(
                rf'\b{re.escape(cliente)}=', regex=True
            )
            sub = df[mask].copy()

            if not sub.empty:
                sub['palavras_por_cliente'] = sub['palavras_por_cliente'].astype(str).map(
                    lambda cell, c=cliente: self._somente_kws_do_cliente(cell, c)
                )

            # split opcional do texto em p1..pN
            if SHEETS_SPLIT_TEXT and 'texto_completo' in sub.columns:
                parts_cols = []

                def split_row_text(s):
                    chunks = _split_in_chunks(s)
                    nonlocal parts_cols
                    while len(parts_cols) < len(chunks):
                        parts_cols.append(f"texto_completo_p{len(parts_cols)+1}")
                    return chunks

                all_chunks = sub['texto_completo'].apply(split_row_text) if not sub.empty else pd.Series([], dtype=object)
                for idx, colname in enumerate(parts_cols):
                    sub[colname] = all_chunks.apply(
                        lambda lst, i=idx: lst[i] if i < len(lst) else ""
                    ) if not all_chunks.empty else []
                sub.drop(columns=['texto_completo'], inplace=True, errors='ignore')

            base_cols = ['data_publicacao', 'titulo', 'fonte', 'url', 'palavras_por_cliente', 'resumo', 'data_coleta']
            if SHEETS_SPLIT_TEXT:
                parts = [c for c in sub.columns if c.startswith('texto_completo_p')]
                parts.sort(key=lambda x: int(x.rsplit('p', 1)[-1]))
                cols = base_cols[:4] + ['palavras_por_cliente', 'resumo'] + parts + ['data_coleta']
            else:
                cols = base_cols[:4] + ['palavras_por_cliente', 'resumo', 'texto_completo', 'data_coleta']

            cols = [c for c in cols if c in sub.columns] if not sub.empty else cols
            sub = sub[cols] if not sub.empty else pd.DataFrame(columns=cols)
            sub = _enforce_sheet_limits(sub, MAX_CELL_CHARS)

            # _get_or_create_ws usa o cache — sem read extra
            ws = self._get_or_create_ws(sh, cliente)

            # Pausa entre clientes para não estourar a cota de writes
            time.sleep(1.2)

            # min_cols era max(20, ...): coluna vazia também ocupa célula no
            # limite do documento, então usa só o que a aba precisa.
            self._ensure_header(ws, cols, min_rows=100, min_cols=len(cols))

            existing_df = self._read_ws_df(ws)
            existing_urls = set()
            url_col_name = None
            for c in existing_df.columns:
                if c.strip().lower() == "url":
                    url_col_name = c
                    break
            if url_col_name:
                existing_urls = set(existing_df[url_col_name].astype(str).tolist())

            if not sub.empty and "url" in sub.columns:
                sub = sub[~sub["url"].astype(str).isin(existing_urls)]

            if sub.empty:
                print(f"✓ Aba {ws.title}: nada novo para inserir.")
                continue

            if ws.row_count < 2:
                _gsheets_with_retry(ws.resize, rows=2)

            if ws.col_count < len(cols):
                _gsheets_with_retry(ws.resize, cols=len(cols))

            # A gravação de UM cliente não pode derrubar a rodada: quando a
            # planilha estourou o limite de células, todos os clientes depois
            # deste ficaram sem notícia, mesmo com a coleta tendo dado certo.
            try:
                _gsheets_with_retry(ws.insert_rows, [[]] * len(sub), row=2,
                                    value_input_option="RAW")

                set_with_dataframe(
                    ws,
                    sub,
                    row=2,
                    col=1,
                    include_index=False,
                    include_column_header=False,
                    resize=False,
                )
            except Exception as e:
                falhas_abas.append((ws.title, str(e)[:160]))
                print(f"✗ Falhou ao gravar na aba {ws.title}: {e}")
                continue
            print(f"✓ Inseridas {len(sub)} linhas no topo da aba: {ws.title}")

            # As novas entram no topo, então o excedente é a cauda antiga.
            if LIMITE_LINHAS_ABA > 0 and ws.row_count > LIMITE_LINHAS_ABA + 1:
                sobra = ws.row_count - (LIMITE_LINHAS_ABA + 1)
                try:
                    _gsheets_with_retry(ws.delete_rows,
                                        LIMITE_LINHAS_ABA + 2, ws.row_count)
                    print(f"  ↳ aparadas {sobra} linhas antigas (teto {LIMITE_LINHAS_ABA})")
                except Exception as e:
                    print(f"  ↳ não consegui aparar a aba {ws.title}: {e}")

        if falhas_abas:
            print("\n" + "=" * 80)
            print(f"{len(falhas_abas)} aba(s) não receberam dados:")
            for _aba, _erro in falhas_abas:
                print(f"  - {_aba}: {_erro}")
            if any("too large" in _e for _, _e in falhas_abas):
                print("\nA planilha bateu o limite de 10 milhões de células do "
                      "Google Sheets. Reduza LIMITE_LINHAS_ABA ou desligue "
                      "SHEETS_SPLIT_TEXT/EXTRACT_BODY, que é o que engorda cada linha.")
            raise SystemExit(1)


# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "").strip()
    if not SPREADSHEET_ID:
        print("Defina SPREADSHEET_ID no ambiente.", file=sys.stderr)
        sys.exit(1)

    tz = int(os.getenv("TZ_OFFSET_HOURS", "-3"))
    extrair = os.getenv("EXTRACT_BODY", "1").strip() in ("1", "true", "True", "yes", "on")

    coletor = ColetorNoticias(
        timezone_offset_hours=tz,
        deduplicar=True,
    )

    coletor.coletar_feeds(
        extrair_texto=extrair,
        apenas_hoje=True,
        pausa_seg=0.2,
    )

    coletor.exportar_por_cliente_para_sheets(
        spreadsheet_id=SPREADSHEET_ID,
        apenas_hoje=True,
    )
