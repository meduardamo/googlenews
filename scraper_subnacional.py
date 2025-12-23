import os
import re
import sys
import json
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from urllib.parse import urlparse

import feedparser
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError


MAX_CELL_CHARS = int(os.getenv("MAX_CELL_CHARS", "47000"))
TZ_OFFSET_HOURS = int(os.getenv("TZ_OFFSET_HOURS", "-3"))

SPREADSHEET_ID = (
    os.getenv("PLANILHA_SUBNACIONAL", "").strip()
    or os.getenv("SPREADSHEET_ID", "").strip()
    or os.getenv("PLANILHA", "").strip()
)

ABA_ENTRADA = os.getenv("ABA_ENTRADA", "deduplicado").strip()
COL_LINK = os.getenv("COL_LINK", "Link").strip()
COL_ESTADO = os.getenv("COL_ESTADO", "Estado").strip()

APENAS_HOJE = os.getenv("APENAS_HOJE", "1").strip().lower() in ("1", "true", "yes", "on")
MAX_PER_ITEM = int(os.getenv("MAX_PER_ITEM", "30"))

PAUSA_ENTRE_ITENS = float(os.getenv("PAUSA_ENTRE_ITENS", "0.25"))
PAUSA_ENTRE_ABAS = float(os.getenv("PAUSA_ENTRE_ABAS", "1.0"))

BATCH_ROWS = int(os.getenv("BATCH_ROWS", "60"))
WRITE_CONSOLIDADO = os.getenv("WRITE_CONSOLIDADO", "0").strip().lower() in ("1", "true", "yes", "on")
ABA_CONSOLIDADO = os.getenv("ABA_CONSOLIDADO", "links_com_estado_monitoramento").strip()

MAX_ESTADOS_POR_RUN = int(os.getenv("MAX_ESTADOS_POR_RUN", "0"))
MAX_MINUTES_BUDGET = int(os.getenv("MAX_MINUTES_BUDGET", "7"))

TOP_RESERVED_ROWS_UF = int(os.getenv("TOP_RESERVED_ROWS_UF", "2"))

# Mantém data_publicacao/data_coleta como data-hora "de verdade" no Sheets (via USER_ENTERED),
# e adiciona colunas só com a data para facilitar ordenação/agrupamento.
OUT_COLS = [
    "data_publicacao",
    "data_publicacao_dia",
    "estado",
    "dominio",
    "titulo",
    "url",
    "data_coleta",
    "data_coleta_dia",
]

DATE_COLS = {"data_publicacao", "data_publicacao_dia", "data_coleta", "data_coleta_dia"}

ESTADOS_SET = {
    "ACRE", "ALAGOAS", "AMAPÁ", "AMAZONAS", "BAHIA", "CEARÁ",
    "DISTRITO FEDERAL", "ESPÍRITO SANTO", "GOIÁS", "MARANHÃO",
    "MATO GROSSO", "MATO GROSSO DO SUL", "MINAS GERAIS", "PARÁ",
    "PARAÍBA", "PARANÁ", "PERNAMBUCO", "PIAUÍ", "RIO DE JANEIRO",
    "RIO GRANDE DO NORTE", "RIO GRANDE DO SUL", "RONDÔNIA",
    "RORAIMA", "SANTA CATARINA", "SÃO PAULO", "SERGIPE", "TOCANTINS"
}

POLITICA_PATTERNS = [
    r"/politica\b",
    r"/poder\b",
    r"/eleicao\b",
    r"/eleicoes\b",
    r"/eleitoral\b",
    r"/campanha\b",
    r"/governo\b",
    r"/congresso\b",
    r"/assembleia\b",
    r"/camara\b",
    r"/senado\b",
]
POLITICA_RE = re.compile("|".join(POLITICA_PATTERNS), flags=re.I)

TITLE_KEYWORDS_RE = re.compile(
    r"\b(politic|poder|eleiç|eleic|govern|prefeit|câmara|camara|senad|deputad|vereador|assembleia)\b",
    flags=re.I,
)


def _now_local():
    return datetime.now() + timedelta(hours=TZ_OFFSET_HOURS)


def _clean_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    return "".join(ch for ch in s if (ch == "\n") or ((ord(ch) >= 32) and (ord(ch) != 127)))


def _truncate(s: str, limit: int = MAX_CELL_CHARS) -> str:
    s = _clean_text(s)
    if len(s) <= limit:
        return s
    return s[: max(0, limit - 10)] + " [..]"


def _safe_user_entered_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = _clean_text(s)
    if s.lstrip().startswith("="):
        return "'" + s
    return s


def _normalize_estado(estado: str) -> str:
    return re.sub(r"\s+", " ", (estado or "").strip()).upper()


def _normalize_url(u: str) -> str:
    if not u:
        return ""
    u = u.strip()
    if not re.match(r"^https?://", u, flags=re.I):
        u = "https://" + u
    return u


def _extract_domain(url: str) -> str:
    if not url:
        return ""
    u = _normalize_url(url)
    try:
        netloc = urlparse(u).netloc.lower()
        netloc = netloc.split("@")[-1]
        netloc = netloc.split(":")[0]
        if netloc.startswith("www."):
            netloc = netloc[4:]
        return netloc
    except Exception:
        return ""


def _domain_and_prefix_from_link(link: str):
    u = _normalize_url(link)
    p = urlparse(u)
    domain = _extract_domain(u)

    path = (p.path or "").strip()
    if not path or path == "/":
        return domain, None

    prefix = f"{p.scheme}://{p.netloc}{path}"
    prefix = prefix.rstrip("/") + "/"
    return domain, prefix


def _google_news_rss_for_query(q: str) -> str:
    q = q.strip().replace(" ", "+")
    return f"https://news.google.com/rss/search?q={q}&hl=pt-BR&gl=BR&ceid=BR:pt-419"


def _google_news_query(domain: str) -> str:
    base = f"site:{domain}"
    filtro = "(politica OR poder OR eleicao OR eleicoes OR eleitoral OR governo OR congresso)"
    return f"{base} {filtro}"


def _entry_datetime(entry):
    dt = None
    if hasattr(entry, "published_parsed") and entry.published_parsed:
        dt = datetime(*entry.published_parsed[:6])
    elif hasattr(entry, "updated_parsed") and entry.updated_parsed:
        dt = datetime(*entry.updated_parsed[:6])
    elif "published" in entry:
        dt2 = pd.to_datetime(entry.get("published"), errors="coerce")
        dt = None if pd.isna(dt2) else dt2.to_pydatetime()
    elif "updated" in entry:
        dt2 = pd.to_datetime(entry.get("updated"), errors="coerce")
        dt = None if pd.isna(dt2) else dt2.to_pydatetime()

    if not dt:
        return None
    return dt + timedelta(hours=TZ_OFFSET_HOURS)


def _eh_hoje(dt_obj: datetime) -> bool:
    return bool(dt_obj) and dt_obj.date() == _now_local().date()


def _fmt_dt(dt_obj: datetime | None) -> str:
    if not dt_obj:
        return ""
    # dd/mm/yyyy hh:mm:ss tende a virar datetime nativo em planilha pt-BR quando USER_ENTERED
    return dt_obj.strftime("%d/%m/%Y %H:%M:%S")


def _fmt_date(dt_obj: datetime | None) -> str:
    if not dt_obj:
        return ""
    return dt_obj.strftime("%d/%m/%Y")


def _chunk_list(lst, n):
    return [lst[i: i + n] for i in range(0, len(lst), n)]


def _gsheets_client_from_env():
    sa_json = os.getenv("GCP_SERVICE_ACCOUNT_JSON", "").strip()
    if not sa_json:
        raise RuntimeError("Faltou o secret GCP_SERVICE_ACCOUNT_JSON no ambiente.")
    info = json.loads(sa_json)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)


def _is_retryable_apierror(ex: Exception) -> bool:
    if not isinstance(ex, APIError):
        return False
    try:
        code = ex.response.status_code
        return code in (429, 500, 502, 503, 504)
    except Exception:
        return False


def _call_with_backoff(fn, *, what: str, max_tries: int = 8, base_sleep: float = 2.0):
    tries = 0
    while True:
        tries += 1
        try:
            return fn()
        except Exception as ex:
            if tries >= max_tries or not _is_retryable_apierror(ex):
                raise
            sleep_s = min(base_sleep * (2 ** (tries - 1)), 90.0)
            print(f"Quota/instabilidade em '{what}' (tentativa {tries}/{max_tries}). Dormindo {sleep_s:.1f}s...")
            time.sleep(sleep_s)


def _read_ws_values(ws):
    return _call_with_backoff(lambda: ws.get_all_values(), what=f"get_all_values:{ws.title}")


def _layout_params_for_ws_title(title: str):
    # fallback caso não ache header via varredura
    t = _normalize_estado(title)
    if t in ESTADOS_SET:
        header_row = 1 + TOP_RESERVED_ROWS_UF
        insert_at_row = header_row + 1
        return header_row, insert_at_row
    return 1, 2


def _find_header_row(values, expected_cols_lower: set, search_rows: int = 25):
    if not values:
        return None
    upto = min(len(values), search_rows)
    for i in range(upto):
        row = values[i]
        row_lower = {str(x).strip().lower() for x in row if str(x).strip()}
        if expected_cols_lower.issubset(row_lower):
            return i + 1
    return None


def _get_header_row_and_insert_at(ws, columns) -> tuple[int, int]:
    values = _read_ws_values(ws)
    expected = {c.strip().lower() for c in columns}

    header_row = _find_header_row(values, expected_cols_lower=expected, search_rows=25)
    if header_row:
        return header_row, header_row + 1

    header_row_guess, insert_at_guess = _layout_params_for_ws_title(ws.title)
    return header_row_guess, insert_at_guess


def _read_ws_df(ws):
    values = _read_ws_values(ws)
    if not values:
        return pd.DataFrame()

    expected = {c.strip().lower() for c in OUT_COLS}
    header_row_guess, _ = _layout_params_for_ws_title(ws.title)
    header_row = _find_header_row(values, expected_cols_lower=expected, search_rows=25) or header_row_guess

    if len(values) < header_row:
        return pd.DataFrame()

    header = values[header_row - 1]
    if not header or not any(str(h).strip() for h in header):
        return pd.DataFrame()

    rows = values[header_row:]
    width = len(header)
    norm_rows = [r + [""] * (width - len(r)) for r in rows]
    return pd.DataFrame(norm_rows, columns=[str(h).strip() for h in header])


def _ensure_header(ws, columns):
    values = _read_ws_values(ws)
    expected = {c.strip().lower() for c in columns}

    header_row_guess, _ = _layout_params_for_ws_title(ws.title)
    header_row = _find_header_row(values, expected_cols_lower=expected, search_rows=25)
    if header_row:
        return False

    a1 = gspread.utils.rowcol_to_a1(header_row_guess, 1)
    b1 = gspread.utils.rowcol_to_a1(header_row_guess, len(columns))
    rng = f"{a1}:{b1}"

    _call_with_backoff(
        lambda: ws.update(rng, [columns], value_input_option="RAW"),
        what=f"write_header:{ws.title}",
    )
    return True


def _get_or_create_ws(sh, title, rows=200, cols=30):
    name = (title or "SEM_NOME").strip()[:31]
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        return _call_with_backoff(lambda: sh.add_worksheet(title=name, rows=rows, cols=cols), what=f"add_ws:{name}")


def _dedup_existing_urls(ws) -> set:
    df = _read_ws_df(ws)
    if df.empty:
        return set()

    url_col = next((c for c in df.columns if c.strip().lower() == "url"), None)
    if not url_col:
        return set()

    return set(df[url_col].astype(str).tolist())


def _insert_rows_once(ws, n_rows: int, insert_at_row: int):
    if n_rows <= 0:
        return
    body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": ws.id,
                        "dimension": "ROWS",
                        "startIndex": insert_at_row - 1,
                        "endIndex": (insert_at_row - 1) + n_rows,
                    },
                    "inheritFromBefore": False,
                }
            }
        ]
    }
    _call_with_backoff(lambda: ws.spreadsheet.batch_update(body), what=f"insertDimension:{ws.title}")


def _write_values(ws, start_row: int, start_col: int, values_2d: list):
    if not values_2d:
        return
    end_row = start_row + len(values_2d) - 1
    end_col = start_col + len(values_2d[0]) - 1
    a1 = gspread.utils.rowcol_to_a1(start_row, start_col)
    b1 = gspread.utils.rowcol_to_a1(end_row, end_col)
    rng = f"{a1}:{b1}"

    # USER_ENTERED faz o Sheets interpretar dd/mm/yyyy e dd/mm/yyyy hh:mm:ss como data/data-hora
    _call_with_backoff(
        lambda: ws.update(rng, values_2d, value_input_option="USER_ENTERED"),
        what=f"update_values:{ws.title}",
    )


def _write_df_in_top(ws, df_to_write: pd.DataFrame, insert_at_row: int) -> int:
    if df_to_write.empty:
        return 0

    df_to_write = df_to_write.copy()

    for col in df_to_write.columns:
        if col in DATE_COLS:
            df_to_write[col] = df_to_write[col].astype(str).map(_truncate)
        else:
            df_to_write[col] = df_to_write[col].astype(str).map(_safe_user_entered_text).map(_truncate)

    values = df_to_write.values.tolist()
    _insert_rows_once(ws, len(values), insert_at_row=insert_at_row)

    chunks = _chunk_list(values, BATCH_ROWS)
    row_cursor = insert_at_row
    total = 0

    for block in chunks:
        _write_values(ws, row_cursor, 1, block)
        total += len(block)
        row_cursor += len(block)
        time.sleep(1.6)

    return total


def _passa_filtro_politica(url: str, titulo: str, prefix: str | None) -> bool:
    url = (url or "").strip()
    titulo = (titulo or "").strip()

    if prefix:
        if not url.startswith(prefix):
            return False
        return True

    if POLITICA_RE.search(url):
        return True

    if TITLE_KEYWORDS_RE.search(titulo):
        return True

    return False


@dataclass
class ItemInput:
    link: str
    estado: str
    dominio: str
    prefix: str | None


def ler_inputs(spreadsheet_id: str) -> list:
    gc = _gsheets_client_from_env()
    sh = gc.open_by_key(spreadsheet_id)

    ws_in = sh.worksheet(ABA_ENTRADA)
    df_in = _read_ws_df(ws_in)
    if df_in.empty:
        raise ValueError(f"A aba '{ABA_ENTRADA}' não trouxe dados (header vazio ou sem linhas).")

    cols_norm = {c.strip().lower(): c for c in df_in.columns}
    if COL_LINK.strip().lower() not in cols_norm or COL_ESTADO.strip().lower() not in cols_norm:
        raise ValueError(
            f"Colunas esperadas não encontradas. Precisa de '{COL_LINK}' e '{COL_ESTADO}'. "
            f"Colunas atuais: {list(df_in.columns)}"
        )

    c_link = cols_norm[COL_LINK.strip().lower()]
    c_estado = cols_norm[COL_ESTADO.strip().lower()]

    out = []
    for _, row in df_in.iterrows():
        link = str(row.get(c_link, "")).strip()
        estado = str(row.get(c_estado, "")).strip()
        if not link or not estado:
            continue

        dom, prefix = _domain_and_prefix_from_link(link)
        if not dom:
            continue

        out.append(ItemInput(link=link, estado=_normalize_estado(estado), dominio=dom, prefix=prefix))

    uniq = {}
    for it in out:
        uniq[(it.estado, it.dominio, it.prefix or "")] = it

    return list(uniq.values())


def coletar_google_news(inputs: list) -> pd.DataFrame:
    rows = []
    seen_urls = set()

    by_state = {}
    for it in inputs:
        by_state.setdefault(it.estado, []).append(it)

    estados = sorted(by_state.keys())
    if MAX_ESTADOS_POR_RUN and MAX_ESTADOS_POR_RUN > 0:
        estados = estados[:MAX_ESTADOS_POR_RUN]

    start = time.time()
    budget_s = max(60, MAX_MINUTES_BUDGET * 60)

    for estado in estados:
        itens = by_state.get(estado, [])
        print(f"\nEstado: {estado} | itens: {len(itens)}")

        for it in itens:
            if (time.time() - start) > budget_s:
                print("Atingiu orçamento de tempo do job, encerrando para evitar falha no Actions.")
                return pd.DataFrame(rows, columns=OUT_COLS)

            q = _google_news_query(it.dominio)
            rss = _google_news_rss_for_query(q)
            feed = feedparser.parse(rss)

            entries = feed.entries[:MAX_PER_ITEM] if feed.entries else []
            print(f"  {it.dominio} -> {len(entries)} entradas (GN RSS)")

            for e in entries:
                titulo = _truncate(e.get("title", "") or "")
                url = (e.get("link", "") or "").strip()
                if not url:
                    continue

                if url in seen_urls:
                    continue

                dt_pub = _entry_datetime(e)
                if APENAS_HOJE and not _eh_hoje(dt_pub):
                    continue

                if not _passa_filtro_politica(url, titulo, it.prefix):
                    continue

                seen_urls.add(url)

                dt_col = _now_local()

                rows.append(
                    {
                        "data_publicacao": _fmt_dt(dt_pub),
                        "data_publicacao_dia": _fmt_date(dt_pub),
                        "estado": estado,
                        "dominio": it.dominio,
                        "titulo": titulo,
                        "url": _truncate(url),
                        "data_coleta": _fmt_dt(dt_col),
                        "data_coleta_dia": _fmt_date(dt_col),
                    }
                )

            if PAUSA_ENTRE_ITENS:
                time.sleep(PAUSA_ENTRE_ITENS)

    df = pd.DataFrame(rows, columns=OUT_COLS)
    for c in OUT_COLS:
        if c in df.columns:
            df[c] = df[c].astype(str).map(_truncate)
    return df


def gravar_por_estado(spreadsheet_id: str, df: pd.DataFrame):
    if df.empty:
        print("Sem notícias para gravar.")
        return

    gc = _gsheets_client_from_env()
    sh = gc.open_by_key(spreadsheet_id)

    if WRITE_CONSOLIDADO:
        ws_all = _get_or_create_ws(sh, ABA_CONSOLIDADO, rows=1200, cols=30)
        _ensure_header(ws_all, OUT_COLS)

        existing_urls_all = _dedup_existing_urls(ws_all)
        df_all = df[~df["url"].astype(str).isin(existing_urls_all)].copy()

        # Insere sempre logo abaixo do header que estiver na planilha (ex.: linha 3 -> insere na 4)
        _, insert_at_row_all = _get_header_row_and_insert_at(ws_all, OUT_COLS)

        if not df_all.empty:
            n = _write_df_in_top(ws_all, df_all[OUT_COLS], insert_at_row=insert_at_row_all)
            print(f"✓ Consolidado: inseridas {n} linhas em {ws_all.title} (a partir da linha {insert_at_row_all}).")
        else:
            print(f"✓ Consolidado: nada novo em {ws_all.title}")

        if PAUSA_ENTRE_ABAS:
            time.sleep(PAUSA_ENTRE_ABAS)

    for estado, sub in df.groupby("estado"):
        estado_tab = _normalize_estado(estado)
        ws = _get_or_create_ws(sh, estado_tab, rows=800, cols=30)

        _ensure_header(ws, OUT_COLS)

        existing_urls = _dedup_existing_urls(ws)
        sub2 = sub[~sub["url"].astype(str).isin(existing_urls)].copy()

        if sub2.empty:
            print(f"✓ {estado_tab}: nada novo.")
            continue

        # Insere sempre logo abaixo do header detectado (ex.: header na linha 3 -> entra na linha 4)
        _, insert_at_row = _get_header_row_and_insert_at(ws, OUT_COLS)

        try:
            n = _write_df_in_top(ws, sub2[OUT_COLS], insert_at_row=insert_at_row)
            print(f"✓ {estado_tab}: inseridas {n} linhas no topo (a partir da linha {insert_at_row}).")
        except APIError as ex:
            if _is_retryable_apierror(ex):
                print(f"Quota estourou ao gravar {estado_tab}. Parando aqui para tentar no próximo ciclo.")
                return
            raise

        if PAUSA_ENTRE_ABAS:
            time.sleep(PAUSA_ENTRE_ABAS)


def coletar_por_uf(spreadsheet_id: str):
    inputs = ler_inputs(spreadsheet_id)
    if not inputs:
        print("Nenhum input válido encontrado (Link/Estado).")
        return

    df = coletar_google_news(inputs)
    gravar_por_estado(spreadsheet_id, df)


def main():
    if not SPREADSHEET_ID:
        print("Defina PLANILHA_SUBNACIONAL (ou SPREADSHEET_ID/PLANILHA) no ambiente.", file=sys.stderr)
        sys.exit(1)

    coletar_por_uf(SPREADSHEET_ID)


if __name__ == "__main__":
    main()
