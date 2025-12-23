import os
import re
import sys
import json
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from urllib.parse import urlparse, urlunparse, urlencode, quote

import feedparser
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError


# Config
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
MAX_PER_SOURCE = int(os.getenv("MAX_PER_SOURCE", os.getenv("MAX_PER_DOMAIN", "30")))
PAUSA_ENTRE_FONTES = float(os.getenv("PAUSA_ENTRE_FONTES", os.getenv("PAUSA_ENTRE_DOMINIOS", "0.15")))
PAUSA_BETWEEN_WS = float(os.getenv("PAUSA_BETWEEN_WS", "0.8"))

BATCH_ROWS = int(os.getenv("BATCH_ROWS", "60"))
MAX_ESTADOS_POR_RUN = int(os.getenv("MAX_ESTADOS_POR_RUN", "10"))
MAX_MINUTES_BUDGET = int(os.getenv("MAX_MINUTES_BUDGET", "6"))

WRITE_CONSOLIDADO = os.getenv("WRITE_CONSOLIDADO", "0").strip().lower() in ("1", "true", "yes", "on")
ABA_CONSOLIDADO = os.getenv("ABA_CONSOLIDADO", "links_com_estado_monitoramento").strip()

OUT_COLS = ["data_publicacao", "estado", "dominio", "titulo", "url", "data_coleta"]

# Layout das abas UF: linha 1 = botão voltar; linha 2 = respiro; linha 3 = header; linha 4+ = dados
TOP_RESERVED_ROWS_UF = int(os.getenv("TOP_RESERVED_ROWS_UF", "2"))

INURL_TERMS = ("politica", "eleicoes", "eleicao", "poder")

ESTADOS_SET = {
    "ACRE", "ALAGOAS", "AMAPÁ", "AMAZONAS", "BAHIA", "CEARÁ",
    "DISTRITO FEDERAL", "ESPÍRITO SANTO", "GOIÁS", "MARANHÃO",
    "MATO GROSSO", "MATO GROSSO DO SUL", "MINAS GERAIS", "PARÁ",
    "PARAÍBA", "PARANÁ", "PERNAMBUCO", "PIAUÍ", "RIO DE JANEIRO",
    "RIO GRANDE DO NORTE", "RIO GRANDE DO SUL", "RONDÔNIA",
    "RORAIMA", "SANTA CATARINA", "SÃO PAULO", "SERGIPE", "TOCANTINS"
}


# Utils
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


def _normalize_estado(estado: str) -> str:
    return re.sub(r"\s+", " ", (estado or "").strip()).upper()


def _ensure_http(url: str) -> str:
    if not url:
        return ""
    u = url.strip()
    if not re.match(r"^https?://", u, flags=re.I):
        u = "https://" + u
    return u


def _ensure_trailing_slash(u: str) -> str:
    return u if u.endswith("/") else (u + "/")


def _extract_domain(url: str) -> str:
    if not url:
        return ""
    u = _ensure_http(url)
    try:
        netloc = urlparse(u).netloc.lower()
        netloc = netloc.split("@")[-1]
        netloc = netloc.split(":")[0]
        if netloc.startswith("www."):
            netloc = netloc[4:]
        return netloc
    except Exception:
        return ""


def _path_hint_from_link(link: str) -> str:
    u = _ensure_http(link)
    try:
        p = urlparse(u)
        path = (p.path or "").strip("/").lower()
        segs = [s for s in path.split("/") if s]
        if not segs:
            return ""
        drop = {"category", "categoria", "noticias", "news", "editoria", "editorias", "tags", "tag"}
        segs2 = [s for s in segs if s not in drop]
        if not segs2:
            segs2 = segs
        return segs2[-1]
    except Exception:
        return ""


def _base_site_url(link: str) -> str:
    u = _ensure_http(link)
    p = urlparse(u)
    return urlunparse((p.scheme, p.netloc, "/", "", "", ""))


def _safe_add_query(url: str, extra: dict) -> str:
    u = _ensure_http(url)
    p = urlparse(u)
    existing = {}
    if p.query:
        for kv in p.query.split("&"):
            if "=" in kv:
                k, v = kv.split("=", 1)
                existing[k] = v
    existing.update(extra)
    new_q = urlencode(existing)
    return urlunparse((p.scheme, p.netloc, p.path, p.params, new_q, p.fragment))


def _candidate_feed_urls_from_section(section_url: str) -> list:
    u = _ensure_http(section_url)
    if not u:
        return []

    low = u.lower()

    # Se já aponta para algo "feed-like", prioriza isso
    if any(x in low for x in ("/feed", "feed=", "/rss", ".rss", ".xml", "atom")):
        return [u]

    u = _ensure_trailing_slash(u)
    base = u.rstrip("/")

    candidates = []

    # WordPress / padrões comuns de seção
    candidates.append(u + "feed/")
    candidates.append(base + "/feed")
    candidates.append(_safe_add_query(u, {"feed": "rss2"}))
    candidates.append(_safe_add_query(u, {"feed": "rss"}))
    candidates.append(_safe_add_query(u, {"output": "rss"}))
    candidates.append(_safe_add_query(u, {"format": "rss"}))
    candidates.append(u + "rss/")
    candidates.append(base + "/rss")
    candidates.append(u + "rss.xml")
    candidates.append(u + "index.xml")
    candidates.append(u + "atom.xml")

    # dedup preservando ordem
    seen = set()
    out = []
    for c in candidates:
        c = c.strip()
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _candidate_feed_urls_from_site_root(section_url: str) -> list:
    site = _base_site_url(section_url)
    site = _ensure_trailing_slash(site)

    candidates = [
        site + "feed/",
        site.rstrip("/") + "/feed",
        _safe_add_query(site, {"feed": "rss2"}),
        _safe_add_query(site, {"feed": "rss"}),
        site + "rss/",
        site.rstrip("/") + "/rss",
        site + "rss.xml",
        site + "index.xml",
        site + "atom.xml",
    ]

    seen = set()
    out = []
    for c in candidates:
        c = c.strip()
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _google_news_rss_for_link(link: str) -> str:
    domain = _extract_domain(link)
    hint = _path_hint_from_link(link)

    inurl_block = " OR ".join([f"inurl:{t}" for t in INURL_TERMS])
    parts = [f"site:{domain}", f"({inurl_block})"]

    if hint:
        parts.append(f'("{domain}/{hint}")')

    q = " ".join(parts)
    q_enc = quote(q, safe="")
    return f"https://news.google.com/rss/search?q={q_enc}&hl=pt-BR&gl=BR&ceid=BR:pt-419"


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
    return bool(dt_obj) and dt_obj.date() == date.today()


def _chunk_list(lst, n):
    return [lst[i: i + n] for i in range(0, len(lst), n)]


# Google Sheets auth
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
            sleep_s = base_sleep * (2 ** (tries - 1))
            sleep_s = min(sleep_s, 60.0)
            print(f"Quota/instabilidade em '{what}' (tentativa {tries}/{max_tries}). Dormindo {sleep_s:.1f}s...")
            time.sleep(sleep_s)


# Sheet helpers
def _read_ws_values(ws):
    return _call_with_backoff(lambda: ws.get_all_values(), what=f"get_all_values:{ws.title}")


def _layout_params_for_ws_title(title: str):
    t = _normalize_estado(title)
    if t in ESTADOS_SET:
        header_row = 1 + TOP_RESERVED_ROWS_UF
        insert_at_row = header_row + 1
        return header_row, insert_at_row
    return 1, 2


def _find_header_row(values, expected_cols_lower: set, search_rows: int = 20):
    if not values:
        return None
    upto = min(len(values), search_rows)
    for i in range(upto):
        row = values[i]
        row_lower = {str(x).strip().lower() for x in row if str(x).strip()}
        if expected_cols_lower.issubset(row_lower):
            return i + 1
    return None


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
        what=f"write_header:{ws.title}"
    )
    return True


def _get_or_create_ws(sh, title, rows=200, cols=30):
    name = (title or "SEM_NOME").strip()[:31]
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        return _call_with_backoff(lambda: sh.add_worksheet(title=name, rows=rows, cols=cols), what=f"add_ws:{name}")


def _dedup_existing_urls(ws) -> set:
    existing = _read_ws_df(ws)
    if existing.empty:
        return set()

    url_col = next((c for c in existing.columns if c.strip().lower() == "url"), None)
    if not url_col:
        return set()

    return set(existing[url_col].astype(str).tolist())


def _insert_rows_top_batch(ws, n_rows: int, insert_at_row: int):
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
    end_row = start_row + len(values_2d) - 1
    end_col = start_col + (len(values_2d[0]) if values_2d else 0) - 1
    a1 = gspread.utils.rowcol_to_a1(start_row, start_col)
    b1 = gspread.utils.rowcol_to_a1(end_row, end_col)
    rng = f"{a1}:{b1}"
    _call_with_backoff(lambda: ws.update(rng, values_2d, value_input_option="RAW"), what=f"update_values:{ws.title}")


def _write_df_in_top(ws, df_to_write: pd.DataFrame, insert_at_row: int) -> int:
    if df_to_write.empty:
        return 0

    df_to_write = df_to_write.copy()
    for col in df_to_write.columns:
        df_to_write[col] = df_to_write[col].astype(str).map(_truncate)

    total_rows = 0
    chunks = _chunk_list(df_to_write.values.tolist(), BATCH_ROWS)

    for values_2d in chunks:
        _insert_rows_top_batch(ws, len(values_2d), insert_at_row=insert_at_row)
        _write_values(ws, insert_at_row, 1, values_2d)
        total_rows += len(values_2d)
        time.sleep(1.2)

    return total_rows


# Input model
@dataclass
class ItemInput:
    link: str
    estado: str
    section_feed_candidates: list
    site_feed_candidates: list
    gn_rss: str


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

        link_http = _ensure_http(link)
        if not link_http:
            continue

        section_candidates = _candidate_feed_urls_from_section(link_http)
        site_candidates = _candidate_feed_urls_from_site_root(link_http)
        gn_rss = _google_news_rss_for_link(link_http)

        out.append(
            ItemInput(
                link=link_http,
                estado=_normalize_estado(estado),
                section_feed_candidates=section_candidates,
                site_feed_candidates=site_candidates,
                gn_rss=gn_rss
            )
        )

    # dedup por UF + Link
    uniq = {}
    for it in out:
        uniq[(it.estado, it.link)] = it

    return list(uniq.values())


# Feed collection
def _parse_first_working_feed(urls: list) -> tuple:
    # Retorna (entries, used_url) ou ([], "")
    for u in urls:
        try:
            feed = feedparser.parse(u)
            entries = feed.entries if hasattr(feed, "entries") else []
            # "bozo" indica feed quebrado; mas alguns ainda vêm com entries
            if entries:
                return entries, u
        except Exception:
            continue
    return [], ""


def coletar_por_links(inputs: list) -> pd.DataFrame:
    rows = []
    seen_urls = set()

    by_state = {}
    for it in inputs:
        by_state.setdefault(it.estado, []).append(it)

    estados = sorted(by_state.keys())
    if MAX_ESTADOS_POR_RUN > 0:
        estados = estados[:MAX_ESTADOS_POR_RUN]

    start = time.time()
    budget_s = max(60, MAX_MINUTES_BUDGET * 60)

    for estado in estados:
        sources = by_state.get(estado, [])
        print(f"\nEstado: {estado} | fontes: {len(sources)}")

        for src in sources:
            if (time.time() - start) > budget_s:
                print("Atingiu orçamento de tempo do job, encerrando coleta para evitar falha.")
                return pd.DataFrame(rows, columns=OUT_COLS)

            # 1) feed da seção
            entries, used = _parse_first_working_feed(src.section_feed_candidates)

            # 2) feed da raiz do site
            if not entries:
                entries, used = _parse_first_working_feed(src.site_feed_candidates)

            # 3) fallback GN com inurl(politica|eleicoes|eleicao)
            feed_source = ""
            if entries:
                feed_source = used
                print(f"  feed OK -> {min(len(entries), MAX_PER_SOURCE)} entradas | {used}")
            else:
                feed = feedparser.parse(src.gn_rss)
                entries = feed.entries[:MAX_PER_SOURCE] if feed.entries else []
                feed_source = "GN"
                print(f"  fallback GN -> {len(entries)} entradas | {src.link}")

            entries = entries[:MAX_PER_SOURCE] if entries else []

            for e in entries:
                titulo = _truncate(e.get("title", "") or "")
                url = (e.get("link", "") or "").strip()
                if not url:
                    continue
                if url in seen_urls:
                    continue
                seen_urls.add(url)

                dt_pub = _entry_datetime(e)

                if APENAS_HOJE and not _eh_hoje(dt_pub):
                    continue

                rows.append(
                    {
                        "data_publicacao": dt_pub.strftime("%Y-%m-%d %H:%M:%S") if dt_pub else "",
                        "estado": estado,
                        "dominio": _extract_domain(url),
                        "titulo": titulo,
                        "url": _truncate(url),
                        "data_coleta": _now_local().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )

            if PAUSA_ENTRE_FONTES:
                time.sleep(PAUSA_ENTRE_FONTES)

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

    # Consolidado opcional
    if WRITE_CONSOLIDADO:
        ws_all = _get_or_create_ws(sh, ABA_CONSOLIDADO, rows=800, cols=30)
        _ensure_header(ws_all, OUT_COLS)

        existing_urls_all = _dedup_existing_urls(ws_all)
        df_all = df[~df["url"].astype(str).isin(existing_urls_all)].copy()

        _, insert_at_row_all = _layout_params_for_ws_title(ws_all.title)

        if not df_all.empty:
            n = _write_df_in_top(ws_all, df_all[OUT_COLS], insert_at_row=insert_at_row_all)
            print(f"✓ Consolidado: inseridas {n} linhas em {ws_all.title}")
        else:
            print(f"✓ Consolidado: nada novo em {ws_all.title}")

        if PAUSA_BETWEEN_WS:
            time.sleep(PAUSA_BETWEEN_WS)

    # Por UF
    for estado, sub in df.groupby("estado"):
        estado_tab = _normalize_estado(estado)
        ws = _get_or_create_ws(sh, estado_tab, rows=400, cols=30)

        _ensure_header(ws, OUT_COLS)

        existing_urls = _dedup_existing_urls(ws)
        sub2 = sub[~sub["url"].astype(str).isin(existing_urls)].copy()

        if sub2.empty:
            print(f"✓ {estado_tab}: nada novo.")
            continue

        _, insert_at_row = _layout_params_for_ws_title(ws.title)

        try:
            n = _write_df_in_top(ws, sub2[OUT_COLS], insert_at_row=insert_at_row)
            print(f"✓ {estado_tab}: inseridas {n} linhas no topo (linha {insert_at_row}).")
        except APIError as ex:
            if _is_retryable_apierror(ex):
                print(f"Quota estourou ao gravar {estado_tab}. Parando para tentar no próximo ciclo.")
                return
            raise

        if PAUSA_BETWEEN_WS:
            time.sleep(PAUSA_BETWEEN_WS)


def coletar_por_uf(spreadsheet_id: str):
    inputs = ler_inputs(spreadsheet_id)
    if not inputs:
        print("Nenhum input válido encontrado (Link/Estado).")
        return

    df = coletar_por_links(inputs)
    gravar_por_estado(spreadsheet_id, df)


def main():
    if not SPREADSHEET_ID:
        print("Defina PLANILHA_SUBNACIONAL (ou SPREADSHEET_ID/PLANILHA) no ambiente.", file=sys.stderr)
        sys.exit(1)

    coletar_por_uf(SPREADSHEET_ID)


if __name__ == "__main__":
    main()
