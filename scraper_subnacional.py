import os
import re
import sys
import json
import time
from datetime import datetime, timedelta, date
from dataclasses import dataclass
from urllib.parse import urlparse

import feedparser
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError


MAX_CELL_CHARS = int(os.getenv("MAX_CELL_CHARS", "47000"))
TZ_OFFSET_HOURS = int(os.getenv("TZ_OFFSET_HOURS", "-3"))

PLANILHA_ENV = os.getenv("PLANILHA_SUBNACIONAL", "").strip() or os.getenv("SPREADSHEET_ID", "").strip()
ABA_ENTRADA = os.getenv("ABA_ENTRADA", "deduplicado").strip()
COL_LINK = os.getenv("COL_LINK", "Link").strip()
COL_ESTADO = os.getenv("COL_ESTADO", "Estado").strip()

APENAS_HOJE = os.getenv("APENAS_HOJE", "1").strip().lower() in ("1", "true", "yes", "on")
MAX_PER_DOMAIN = int(os.getenv("MAX_PER_DOMAIN", "30"))

PAUSA_ENTRE_DOMINIOS = float(os.getenv("PAUSA_ENTRE_DOMINIOS", "0.15"))
PAUSA_BETWEEN_WS = float(os.getenv("PAUSA_BETWEEN_WS", "0.8"))

BATCH_ROWS = int(os.getenv("BATCH_ROWS", "60"))

MAX_ESTADOS_POR_RUN = int(os.getenv("MAX_ESTADOS_POR_RUN", "10"))
MAX_MINUTES_BUDGET = int(os.getenv("MAX_MINUTES_BUDGET", "5"))

WRITE_CONSOLIDADO = os.getenv("WRITE_CONSOLIDADO", "0").strip().lower() in ("1", "true", "yes", "on")
ABA_CONSOLIDADO = os.getenv("ABA_CONSOLIDADO", "links_com_estado_monitoramento").strip()


OUT_COLS = [
    "data_publicacao",
    "estado",
    "dominio",
    "fonte",
    "titulo",
    "url",
    "resumo",
    "data_coleta",
]


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


def _extract_domain(url: str) -> str:
    if not url:
        return ""
    u = url.strip()
    if not re.match(r"^https?://", u, flags=re.I):
        u = "https://" + u
    try:
        netloc = urlparse(u).netloc.lower()
        netloc = netloc.split("@")[-1]
        netloc = netloc.split(":")[0]
        if netloc.startswith("www."):
            netloc = netloc[4:]
        return netloc
    except Exception:
        return ""


def _google_news_rss_for_domain(domain: str) -> str:
    q = f"site:{domain}"
    return f"https://news.google.com/rss/search?q={q}&hl=pt-BR&gl=BR&ceid=BR:pt-419"


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
    return [lst[i : i + n] for i in range(0, len(lst), n)]


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


def _read_ws_values(ws):
    return _call_with_backoff(lambda: ws.get_all_values(), what=f"get_all_values:{ws.title}")


def _read_ws_df(ws):
    values = _read_ws_values(ws)
    if not values:
        return pd.DataFrame()
    header = values[0]
    if not header or not any(h.strip() for h in header):
        return pd.DataFrame()
    rows = values[1:]
    width = len(header)
    norm_rows = [r + [""] * (width - len(r)) for r in rows]
    return pd.DataFrame(norm_rows, columns=[h.strip() for h in header])


def _ensure_header(ws, columns):
    values = _read_ws_values(ws)
    if not values or not values[0] or not any(v.strip() for v in values[0]):
        rng = f"A1:{gspread.utils.rowcol_to_a1(1, len(columns))}"
        _call_with_backoff(lambda: ws.update(rng, [columns], value_input_option="RAW"), what=f"write_header:{ws.title}")
        return True
    return False


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


@dataclass
class ItemInput:
    link: str
    estado: str
    dominio: str


def ler_inputs(spreadsheet_id: str) -> list[ItemInput]:
    gc = _gsheets_client_from_env()
    sh = gc.open_by_key(spreadsheet_id)

    try:
        ws_in = sh.worksheet(ABA_ENTRADA)
    except gspread.WorksheetNotFound:
        raise RuntimeError(f"Aba de entrada '{ABA_ENTRADA}' não existe.")

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
        dom = _extract_domain(link)
        if not dom:
            continue
        out.append(ItemInput(link=link, estado=_normalize_estado(estado), dominio=dom))

    uniq = {}
    for it in out:
        uniq[(it.estado, it.dominio)] = it

    return list(uniq.values())


def coletar_google_news_por_dominios(inputs: list[ItemInput]) -> pd.DataFrame:
    rows = []
    seen = set()

    by_state = {}
    for it in inputs:
        by_state.setdefault(it.estado, []).append(it.dominio)

    estados = sorted(by_state.keys())
    if MAX_ESTADOS_POR_RUN > 0:
        estados = estados[:MAX_ESTADOS_POR_RUN]

    start = time.time()
    budget_s = max(60, MAX_MINUTES_BUDGET * 60)

    for estado in estados:
        dominios = sorted(set(by_state.get(estado, [])))
        print(f"\nEstado: {estado} | domínios: {len(dominios)}")

        for dom in dominios:
            if (time.time() - start) > budget_s:
                print("Atingiu orçamento de tempo do job, encerrando coleta para evitar falha no Actions.")
                df = pd.DataFrame(rows, columns=OUT_COLS)
                return df

            rss = _google_news_rss_for_domain(dom)
            feed = feedparser.parse(rss)

            entries = feed.entries[:MAX_PER_DOMAIN] if feed.entries else []
            print(f"  {dom} -> {len(entries)} entradas (GN RSS)")

            for e in entries:
                titulo = _truncate(e.get("title", "") or "")
                url = (e.get("link", "") or "").strip()
                resumo = _truncate(e.get("summary", "") or "")

                if not url:
                    continue
                if url in seen:
                    continue
                seen.add(url)

                dt_pub = _entry_datetime(e)
                if APENAS_HOJE and not _eh_hoje(dt_pub):
                    continue

                rows.append(
                    {
                        "data_publicacao": dt_pub.strftime("%Y-%m-%d %H:%M:%S") if dt_pub else "",
                        "estado": estado,
                        "dominio": dom,
                        "fonte": "Google News (RSS)",
                        "titulo": titulo,
                        "url": _truncate(url),
                        "resumo": resumo,
                        "data_coleta": _now_local().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )

            if PAUSA_ENTRE_DOMINIOS:
                time.sleep(PAUSA_ENTRE_DOMINIOS)

    df = pd.DataFrame(rows, columns=OUT_COLS)

    for c in ["titulo", "resumo", "url", "estado", "dominio", "fonte", "data_publicacao", "data_coleta"]:
        if c in df.columns:
            df[c] = df[c].astype(str).map(_truncate)

    return df


def _insert_rows_top_batch(ws, n_rows: int, insert_at_row: int = 2):
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


def _write_values(ws, start_row: int, start_col: int, values_2d: list[list[str]]):
    end_row = start_row + len(values_2d) - 1
    end_col = start_col + (len(values_2d[0]) if values_2d else 0) - 1
    a1 = gspread.utils.rowcol_to_a1(start_row, start_col)
    b1 = gspread.utils.rowcol_to_a1(end_row, end_col)
    rng = f"{a1}:{b1}"

    _call_with_backoff(lambda: ws.update(rng, values_2d, value_input_option="RAW"), what=f"update_values:{ws.title}")


def _write_df_in_top(ws, df_to_write: pd.DataFrame, insert_at_row: int = 2) -> int:
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


def gravar_por_estado(spreadsheet_id: str, df: pd.DataFrame):
    if df.empty:
        print("Sem notícias para gravar.")
        return

    gc = _gsheets_client_from_env()
    sh = gc.open_by_key(spreadsheet_id)

    if WRITE_CONSOLIDADO:
        ws_all = _get_or_create_ws(sh, ABA_CONSOLIDADO, rows=500, cols=30)
        _ensure_header(ws_all, OUT_COLS)

        existing_urls_all = _dedup_existing_urls(ws_all)
        df_all = df[~df["url"].astype(str).isin(existing_urls_all)].copy()

        if not df_all.empty:
            n = _write_df_in_top(ws_all, df_all[OUT_COLS], insert_at_row=2)
            print(f"✓ Consolidado: inseridas {n} linhas em {ws_all.title}")
        else:
            print(f"✓ Consolidado: nada novo em {ws_all.title}")

        if PAUSA_BETWEEN_WS:
            time.sleep(PAUSA_BETWEEN_WS)

    for estado, sub in df.groupby("estado"):
        estado_tab = _normalize_estado(estado)
        ws = _get_or_create_ws(sh, estado_tab, rows=200, cols=30)

        _ensure_header(ws, OUT_COLS)

        existing_urls = _dedup_existing_urls(ws)
        sub2 = sub[~sub["url"].astype(str).isin(existing_urls)].copy()

        if sub2.empty:
            print(f"✓ {estado_tab}: nada novo.")
            continue

        try:
            n = _write_df_in_top(ws, sub2[OUT_COLS], insert_at_row=2)
            print(f"✓ {estado_tab}: inseridas {n} linhas no topo.")
        except APIError as ex:
            if _is_retryable_apierror(ex):
                print(f"Quota estourou no meio da execução ao gravar {estado_tab}. Encerrando para tentar no próximo ciclo.")
                return
            raise

        if PAUSA_BETWEEN_WS:
            time.sleep(PAUSA_BETWEEN_WS)


def coletar_por_uf(spreadsheet_id: str):
    inputs = ler_inputs(spreadsheet_id)
    if not inputs:
        print("Nenhum input válido encontrado (Link/Estado).")
        return

    df = coletar_google_news_por_dominios(inputs)
    gravar_por_estado(spreadsheet_id, df)


def main():
    if not PLANILHA_ENV:
        print("Defina PLANILHA_SUBNACIONAL (ou SPREADSHEET_ID) no ambiente.", file=sys.stderr)
        sys.exit(1)

    coletar_por_uf(PLANILHA_ENV)


if __name__ == "__main__":
    main()
