"""Teste do prompt novo de alinhamento: reclassifica uma amostra das linhas já
classificadas e imprime antiga x nova. Não grava nada na planilha."""
import os, random
from string import Template
import alinhamento as A

REGRAS_NOVAS = """
Como decidir (vale antes das classes abaixo):
- Decida pelo ASSUNTO PRINCIPAL do Conteúdo (o que o título e o início do texto tratam). Tema do cliente citado de passagem, em lista ou como exemplo dentro de conteúdo sobre outro assunto (disputa eleitoral, perfil ou agenda de candidato e autoridade, investigação policial, julgamento, economia geral, mercado, cultura, entretenimento, esporte profissional) é "Não Alinha".
- Só alinha conteúdo com dimensão de política pública, regulação, orçamento, decisão de governo ou de Justiça, dado oficial ou debate público sobre um tema do cliente. Caso individual (crime, acidente, morte, celebridade), dica de saúde, de estudo ou de consumo e fato do dia sem essa dimensão é "Não Alinha".
- Fato de outro país só alinha se for política, regulação, decisão judicial ou dado sobre tema do cliente.
- A presença de uma palavra do tema não prova alinhamento: falta de luz por temporal não é política de energia; terremoto que fechou escolas não é política de educação.
- Ato de rotina sobre empresa ou pessoa específica (outorga, autorização, cassação, liberação de usina, afastamento ou dispensa de servidor, concurso, vestibular) é "Não se aplica".
"""

NOVO = Template(A.PROMPT.template.replace(
    "Regras de evidência:", REGRAS_NOVAS.strip() + "\n\nRegras de evidência:", 1
).replace(
    "(b) Cobertura incompleta — o Conteúdo aborda parcialmente o tema do cliente, mas mistura substancialmente outras agendas não relacionadas, de modo que a relevância é real porém limitada.",
    "(b) Cobertura incompleta — o tema do cliente ocupa parte própria e substancial do Conteúdo (não uma menção), ao lado de outras agendas não relacionadas.",
    1))
assert "ASSUNTO PRINCIPAL" in NOVO.template and "parte própria e substancial" in NOVO.template

N = int(os.getenv("N_AMOSTRA", "90"))
random.seed(11)
linhas = []
for ws in A.sh.worksheets():
    if ws.title in A.SKIP_TITLES:
        continue
    df = A.read_sheet_df(ws)
    if df.empty or A.OUT_ALINH_COL not in df.columns:
        continue
    t, r, b = (A.norm_find_col(df, c) for c in (A.TIT_COL, A.RES_COL, A.BODY_COL))
    for i in range(len(df)):
        if df.at[i, A.OUT_ALINH_COL] in ("Alinha", "Parcial"):
            linhas.append((ws.title, df.loc[i], t, r, b))
amostra = random.sample(linhas, min(N, len(linhas)))
mud = 0
for aba, row, t, r, b in amostra:
    _, desc = A.CLIENTE_DESCRICOES.get(aba, (aba, ""))
    conteudo = A.build_content(row, t, r, b).replace("</conteudo>", "</conteudo​>")
    novo = A.call_gemini(NOVO.substitute(descricao=desc, conteudo=conteudo))
    antigo = row[A.OUT_ALINH_COL]
    mud += novo["alinhamento"] != antigo
    print(f"[{aba}] {antigo} -> {novo['alinhamento']} | {str(row[t])[:110]} || {novo['justificativa'][:160]}")
print(f"\nmudaram {mud} de {len(amostra)}")
