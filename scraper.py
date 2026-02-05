# Coleta notícias e grava direto no Google Sheets (uma aba por cliente)
# - Cada aba mostra só as keywords do próprio cliente.
# - Blindagem 50k chars por célula (truncamento) + opção de split do texto em p1..pN.
# - INSERE novas linhas no topo (linha 2), não sobrescreve e não duplica (chave: URL).

import os, re, time, sys
from datetime import datetime, timedelta, date

import feedparser
import pandas as pd
from newspaper import Article

import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials

# Limites e flags
# Margem pra baixo do teto de 50k do Sheets
MAX_CELL_CHARS = int(os.getenv("MAX_CELL_CHARS", "47000"))
# Se 1/true → quebra 'texto_completo' em colunas p1..pN; senão só trunca.
SHEETS_SPLIT_TEXT = os.getenv("SHEETS_SPLIT_TEXT", "0").strip() in ("1","true","True","yes","on")

# Helpers de limpeza/limite
def _clean_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    # remove controles (char < 32) exceto \n; remove DEL (127)
    return "".join(ch for ch in s if (ch == "\n") or ((ord(ch) >= 32) and (ord(ch) != 127)))

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
    # Garante que nenhuma célula (objeto/string) passe do limite
    for c in df.columns:
        if df[c].dtype == object or str(df[c].dtype).startswith("string"):
            df[c] = df[c].astype(str).map(lambda v: _truncate(v, limit))
    return df

# Mapa: Cliente → Tema → Keywords (whole-word)
CLIENT_THEME_DATA = """
IAS|Educação|Matemática; Alfabetização; Alfabetização Matemática; Recomposição de aprendizagem; Plano Nacional de Educação
ISG|Educação|Tempo Integral; PNE; FUNDEB; Ensino Técnico Profissionalizante (EPT); Ensino Médio; Propag; Infraestrutura Escolar; Ensino Fundamental Integral; Albetização Integral
IU|Educação|Recomposição da aprendizagem; Educação em tempo integral; Fundeb; Educação e equidade; PNE; Educação Profissional e Técnologica (EPT)
Reúna|Educação|Matemática; Alfabetização; Alfabetização Matemática; Recomposição de aprendizagem; Plano Nacional de Educação; Emendas parlamentares educação
REMS|Esportes|Esporte e desenvolvimento social; esporte e educação; esporte e equidade; paradesporto
FMCSV|Primeira infância|Criança; Criança Feliz (sempre); alfabetização; creche; Conanda; maternidade; parentalidade; paternidade; primeira infancia; infantil; infancia; FUNDEB; educação básica; PNE; Homeschooling
IEPS|Saúde|SUS; Sistema Único de Saúde; Equidade em saúde; Seguros de saúde; populações vulneráveis; desigualdades sociais; Organização do SUS; políticas públicas em saúde; Governança do SUS; Regionalização em saúde; População negra em saúde; Saúde indígena; Povos originários; Saúde da pessoa idosa; envelhecimento ativo; Atenção Primária; Saúde da criança; Saúde do adolescente; Saúde da mulher; Saúde do homem; Saúde da pessoa com deficiência; Saúde da população LGBTQIA+; Financiamento da saúde; Emendas e orçamento da saúde; emendas parlamentares; Ministério da Saúde; Trabalhadores e profissionais de saúde; Força de trabalho em saúde; política de recursos humanos em saúde; Formação profissional de saúde; Cuidados primários em saúde; Emergências climáticas e ambientais em saúde; mudanças ambientais; adaptação climática; saúde ambiental; políticas climáticas; Vigilância em saúde; epidemiológica; Emergência em saúde; estado de emergência; Saúde suplementar; complementar; privada; planos; seguros; seguradoras; planos populares; Anvisa; ANS; Sandbox regulatório; Cartões e administradoras de benefícios em saúde; Economia solidária em saúde mental; Pessoa em situação de rua; saúde mental; Fiscalização de comunidades terapêuticas; Rede de atenção psicossocial; RAPS; unidades de acolhimento; assistência multiprofissional; centros de convivência; Cannabis; canabidiol; tratamento terapêutico; Desinstitucionalização; manicômios; hospitais de custódia; Saúde mental na infância; adolescência; escolas; comunidades escolares; protagonismo juvenil; Dependência química; vícios; ludopatia; Treinamento; capacitação em saúde mental; Intervenções terapêuticas em saúde mental; Internet e redes sociais na saúde mental; Violência psicológica; Surto psicótico
Manual|Saúde|Ozempic; Wegovy; Mounjaro; Telemedicina; Telessaúde; CBD; Cannabis Medicinal; CFM; Conselho Federal de Medicina; Farmácia Magistral; Medicamentos Manipulados; Minoxidil; Emagrecedores; Retenção de receita de medicamentos; tirzepatida; liraglutida
Mevo|Saúde|Prontuário eletrônico; dispensação eletrônica; telessaúde; assinatura digital; certificado digital; controle sanitário; prescrição por enfermeiros; doenças crônicas; responsabilização de plataformas digitais; regulamentação de marketplaces; segurança cibernética; inteligência artificial; digitalização do SUS; venda e distribuição de medicamentos; venda de medicamentos; Bula digital; Atesta CFM; Sistemas de Controle de farmácia; SNGPC; Farmacêutico Remoto; Medicamentos Isentos de Prescrição (MIPs); RNDS - Rede Nacional de Dados em Saúde; Interoperabilidade; Listas de Substâncias Entorpecentes, Psicotrópicas, Precursoras e outras; Substâncias sob controle especial; Tabela SUS; Saúde digital; SEIDIGI; ICP - Brasil; Farmácia popular; CMED
Umane|Saúde|SUS; Sistema Único de Saúde; Equidade em saúde; populações vulneráveis; desigualdades sociais; Organização do SUS; políticas públicas em saúde; Governança do SUS; Regionalização em saúde; População negra em saúde; Saúde indígena; Povos originários; Saúde da pessoa idosa; envelhecimento ativo; Atenção Primária; Saúde da criança; Saúde do adolescente; Saúde da mulher; Saúde do homem; Saúde da pessoa com deficiência; Saúde da população LGBTQIA+; Financiamento da saúde; Emendas e orçamento da saúde; emendas parlamentares; Ministério da Saúde; Trabalhadores e profissionais de saúde; Força de trabalho em saúde; política de recursos humanos em saúde; Formação profissional de saúde; Cuidados primários em saúde; Emergências climáticas e ambientais em saúde; mudanças ambientais; adaptação climática; saúde ambiental; políticas climáticas; Vigilância em saúde; epidemiológica; Emergência em saúde; estado de emergência; Saúde suplementar; complementar; privada; planos; seguros; seguradoras; planos populares; Anvisa; ANS; Sandbox regulatório; Cartões e administradoras de benefícios em saúde; Conass; Conasems
Cactus|Saúde|Saúde mental; saúde mental para meninas; saúde mental para juventude; saúde mental para mulheres; Rede de atenção psicossocial; RAPS; CAPS; Centro de Apoio Psicossocial; programa saúde na escola; bullying; cyberbullying; eca digital
Vital Strategies|Saúde|Saúde mental; Dados para a saúde; Morte evitável; Doenças crônicas não transmissíveis; Rotulagem de bebidas alcoólicas; Educação em saúde; Bebidas alcoólicas; Imposto seletivo; Rotulagem de alimentos; Alimentos ultraprocessados; Publicidade infantil; Publicidade de alimentos ultraprocessados; Tributação de bebidas alcoólicas; Alíquota de bebidas alcoólicas; Cigarro eletrônico; Controle de tabaco; Violência doméstica; Exposição a fatores de risco; Departamento de Saúde Mental; Hipertensão arterial; Saúde digital; Violência contra crianças; Violência contra mulheres; Feminicídio; COP 30
Coletivo Feminista|Direitos reprodutivos|aborto; nascituro; gestação acima de 22 semanas; interrupção legal da gestação; interrupção da gestação; Resolução 258 Conanda; vida por nascer; vida desde a concepção; criança por nascer; infanticídio; feticídio; assistolia fetal; medicamento abortivo; misoprostol; citotec; cytotec; mifepristona; ventre; assassinato de bebês; luto parental; síndrome pós aborto
IDEC|Saúde|Defesa do consumidor; Ação Civil Pública; Arbitragem; SAC; Concorrência; Reforma Tributária; Ultraprocessados; Doenças crônicas não transmissíveis (DCNTs); Obesidade; Codex Alimentarius; Gordura trans; Adoçantes; edulcorantes; Rotulagem de alimentos; Transgênicos; Organismos geneticamente modificados (OGMs); Marketing e publicidade de alimentos; Comunicação mercadológica; Escolas e alimentação escolar; Bebidas açucaradas; refrigerante; Programa Nacional de Alimentação Escolar (PNAE); Educação Alimentar e Nutricional (EAN); Agrotóxicos; pesticidas; defensivos fitossanitários; Orgânicos; Tributação de alimentos não saudáveis; Desertos alimentares; Desperdício de alimentos; Segurança Alimentar e Nutricional (SAN); Direito Humano à Alimentação; Fome; Sustentabilidade; Mudança climática; Plástico; Gestão de resíduos; Economia circular; Desmatamento; Greenwashing; Energia elétrica; Encargos tarifários; Subsídios na tarifa de energia; Descontos na tarifa de energia; Energia pré-paga; Abertura do mercado de energia para consumidor cativo; Mercado livre de energia; Qualidade do serviço de energia; Tarifa Social de Energia Elétrica; Geração térmica; Combustíveis fósseis; Transição energética; Descarbonização da matriz elétrica; Gases de efeito estufa; Acordo de Paris; Objetivos do Desenvolvimento Sustentável; Reestruturação do setor de energia; Reforma do setor elétrico; Modernização do setor elétrico; Itens de custo da tarifa de energia elétrica; Universalização do acesso à energia; Eficiência energética; Geração distribuída; Carvão mineral; Painel solar; Crédito imobiliário; Crédito consignado; Publicidade de crédito; Cartão de crédito; pagamento de fatura; parcelamento com e sem juros; Cartões pré-pagos; Programas de fidelidade; Cheque especial; Taxa de juros; Contrato de crédito; Endividamento de jovens; Crédito estudantil; Endividamento de idosos; Crédito por meio de aplicativos; Abertura e movimentação de conta bancária; Cobrança de serviços sem autorização; Cadastro Positivo; Contratação de serviços bancários com imposição de seguros e títulos de capitalização; Acessibilidade aos canais de serviços bancários (caixa eletrônico, agências, internet banking e aplicativos móveis); Contratação de pacotes de contas bancárias; Acesso à informação em caso de negativa de crédito; Plano de saúde; Saúde suplementar; Medicamentos isentos de prescrição (MIP); Medicamentos antibióticos; antimicrobianos; Propriedade intelectual; Patentes; Licença compulsória; Preços de medicamentos; Complexo Econômico-Industrial da Saúde; Saúde digital; Prontuário eletrônico; Rede Nacional de Dados em Saúde (RNDS); DATASUS; Proteção de dados pessoais; Telessaúde; Telecomunicações; Internet; TV por assinatura; Serviço de Acesso Condicionado (SeAC); Telefonia móvel; Telefonia fixa; TV digital; Lei Geral de Proteção de Dados (LGPD); Autoridade Nacional de Proteção de Dados (ANPD); Reconhecimento facial; Lei Geral de Telecomunicações; Bens reversíveis; Fundo de Universalização dos Serviços de Telecomunicações (FUST); Provedores de acesso; Franquia de internet; Marco Civil da Internet; Neutralidade de rede; zero rating; Privacidade; Lei de Acesso à Informação; Regulação de plataformas digitais; Desinformação; Fake news; Dados biométricos; Vazamento de dados; Telemarketing; Serviço de Valor Adicionado (SVA)
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
    return re.compile(rf'\b{re.escape(term)}\b', flags=re.IGNORECASE)

def parse_palavras_por_cliente_cell(cell: str):
    # Ex.: "IAS=PNE|Alfabetização; IEPS=SUS"
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
            'data_publicacao','titulo','fonte','url',
            'palavras_por_cliente','resumo','texto_completo','data_coleta'
        ])
        self.clientes = parse_client_theme_data(CLIENT_THEME_DATA)
        self.cliente_patterns = {
            c: [make_wholeword_pattern(k) for k in data['keywords']]
            for c, data in self.clientes.items()
        }

        # feeds habilitados
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
            'CNE (via GN)': 'https://news.google.com/rss/search?q=site:portal.mec.gov.br/conselho-nacional-de-educacao',

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
        }

    # ==== datas ====
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

    # matching
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

    # dedup interno
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

    # extração de corpo
    def extrair_texto_completo(self, url):
        try:
            art = Article(url, language='pt')
            art.download()
            art.parse()
            return art.text
        except Exception:
            return ""

    # coleta
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
                    resumo = e.get('summary', '')

                    dt_pub = self._entry_datetime(e)
                    if apenas_hoje and not self._eh_hoje(dt_pub):
                        continue

                    if extrair_texto:
                        corpo = self.extrair_texto_completo(url)
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

    # Google Sheets
    def _gsheets_client(self):
        creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")
        scopes = ["https://www.googleapis.com/auth/spreadsheets",
                  "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        return gspread.authorize(creds)

    def _get_or_create_ws(self, sh, title):
        name = title[:31]
        try:
            return sh.worksheet(name)
        except gspread.WorksheetNotFound:
            return sh.add_worksheet(title=name, rows=100, cols=20)

    def _somente_kws_do_cliente(self, cell, cliente):
        try:
            d = parse_palavras_por_cliente_cell(cell)
            return "|".join(d.get(cliente, []))
        except Exception:
            return ""

    def _read_ws_df(self, ws):
        values = ws.get_all_values()
        if not values:
            return pd.DataFrame()
        header, rows = values[0], values[1:]
        if not any(h.strip() for h in header):
            return pd.DataFrame()
        # normaliza largura
        width = len(header)
        norm_rows = [r + [""] * (width - len(r)) for r in rows]
        df = pd.DataFrame(norm_rows, columns=[h.strip() for h in header])
        return df

    def _ensure_header(self, ws, columns):
        # se a primeira linha não tiver cabeçalho, escreve
        values = ws.get_all_values()
        if not values or not values[0] or not any(v.strip() for v in values[0]):
            tmp = pd.DataFrame(columns=columns)
            set_with_dataframe(ws, tmp, include_index=False, include_column_header=True, resize=True)
            return True
        return False

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

        # sanitiza + trunca os campos base (título/resumo/fonte/url)
        for col in ['titulo', 'resumo', 'fonte', 'url']:
            if col in df.columns:
                df[col] = df[col].astype(str).map(_truncate)

        # texto_completo: trunca se não for split; se for split, só limpa
        if 'texto_completo' in df.columns:
            if SHEETS_SPLIT_TEXT:
                df['texto_completo'] = df['texto_completo'].astype(str).map(_clean_text)
            else:
                df['texto_completo'] = df['texto_completo'].astype(str).map(_truncate)

        for cliente in self.clientes.keys():
            mask = df['palavras_por_cliente'].astype(str).str.contains(rf'\b{re.escape(cliente)}=', regex=True)
            sub = df[mask].copy()

            # mantém só as keywords do cliente
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

                all_chunks = sub['texto_completo'].apply(split_row_text) if not sub.empty else pd.Series([])
                for idx, colname in enumerate(parts_cols):
                    sub[colname] = all_chunks.apply(lambda lst, i=idx: lst[i] if i < len(lst) else "") if not all_chunks.empty else []
                sub.drop(columns=['texto_completo'], inplace=True, errors='ignore')

            # ordem de colunas
            base_cols = ['data_publicacao','titulo','fonte','url','palavras_por_cliente','resumo','data_coleta']
            if SHEETS_SPLIT_TEXT:
                parts = [c for c in sub.columns if c.startswith('texto_completo_p')]
                parts.sort(key=lambda x: int(x.rsplit('p',1)[-1]))
                cols = base_cols[:4] + ['palavras_por_cliente','resumo'] + parts + ['data_coleta']
            else:
                cols = base_cols[:4] + ['palavras_por_cliente','resumo','texto_completo','data_coleta']

            # garante somente colunas existentes e na ordem
            cols = [c for c in cols if c in sub.columns] if not sub.empty else cols
            sub = sub[cols] if not sub.empty else pd.DataFrame(columns=cols)

            # última barreira: garante que nada passa de 50k
            sub = _enforce_sheet_limits(sub, MAX_CELL_CHARS)

            ws = self._get_or_create_ws(sh, cliente)

            # garante cabeçalho (só na primeira vez que a aba existe)
            self._ensure_header(ws, cols)

            # dedup contra o que JÁ está na planilha (por URL)
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

            # INSERE LINHAS no topo (a partir da linha 2)
            ws.insert_rows([[]] * len(sub), row=2, value_input_option="RAW")

            # escreve os dados a partir de A2, sem header
            set_with_dataframe(
                ws,
                sub,
                row=2,
                col=1,
                include_index=False,
                include_column_header=False,
                resize=False
            )
            print(f"✓ Inseridas {len(sub)} linhas no topo da aba: {ws.title}")

# Main
if __name__ == "__main__":
    SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "").strip()
    if not SPREADSHEET_ID:
        print("Defina SPREADSHEET_ID no ambiente.", file=sys.stderr)
        sys.exit(1)

    tz = int(os.getenv("TZ_OFFSET_HOURS", "-3"))
    extrair = os.getenv("EXTRACT_BODY", "1").strip() in ("1","true","True","yes","on")

    coletor = ColetorNoticias(
        timezone_offset_hours=tz,
        deduplicar=True
    )

    coletor.coletar_feeds(
        extrair_texto=extrair,
        apenas_hoje=True,
        pausa_seg=0.2
    )

    coletor.exportar_por_cliente_para_sheets(
        spreadsheet_id=SPREADSHEET_ID,
        apenas_hoje=True
    )
