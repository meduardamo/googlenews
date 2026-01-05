# googlenews

Pipeline de coleta de notícias via RSS/Google News RSS, cobrindo:
- **clipping geral/nacional** (sites gerais e nacionais)
- **monitoramento subnacional** (por UF e fontes locais)

## O que este repo entrega
- Coleta itens (título, url, veículo/fonte, data de publicação/coleta, resumo quando disponível)
- Normaliza campos (datas, strings, URLs)
- Deduplica para evitar entradas repetidas
- Escreve em Google Sheets (abas específicas para geral e subnacional)

## Estrutura do repositório
- `scraper.py`: coleta geral/nacional
- `scraper_subnacional.py`: coleta por UF (subnacional)
- `alinhamento.py`: (se aplicável) rotinas auxiliares de classificação/organização
- `.github/workflows/main.yml`: execução via GitHub Actions
- `requirements.txt`: dependências Python
