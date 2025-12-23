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

DIAS_RETROATIVOS = int(os.getenv("DIAS_RETROATIVOS", "3"))
APENAS_HOJE = os.getenv("APENAS_HOJE", "0").strip().lower() in ("1", "true", "yes", "on")
MAX_PER_ITEM = int(os.getenv("MAX_PER_ITEM", "30"))

PAUSA_ENTRE_ITENS = float(os.getenv("PAUSA_ENTRE_ITENS", "0.25"))
PAUSA_ENTRE_ABAS = float(os.getenv("PAUSA_ENTRE_ABAS", "1.0"))

BATCH_ROWS = int(os.getenv("BATCH_ROWS", "60"))
WRITE_CONSOLIDADO = os.getenv("WRITE_CONSOLIDADO", "0").strip().lower() in ("1", "true", "yes", "on")
ABA_CONSOLIDADO = os.getenv("ABA_CONSOLIDADO", "links_com_estado_monitoramento").strip()

MAX_ESTADOS_POR_RUN = int(os.getenv("MAX_ESTADOS_POR_RUN", "0"))
MAX_MINUTES_BUDGET = int(os.getenv("MAX_MINUTES_BUDGET", "7"))

TOP_RESERVED_ROWS_UF = int(os.getenv("TOP_RESERVED_ROWS_UF", "2"))

OUT_COLS = [
    "data_publicacao",
    "estado",
    "dominio",
    "titulo",
    "url",
    "data_coleta",
]

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


def _normalize_estado(estado: str) -> str:
    return re.sub(r"\s+", " ", (estado or "").strip()).upper()


def _normalize_url(u: str) -> str:
    if not u:
        return ""
    u = u.strip()
    return u


def _domain_from_url(u: str) -> str:
    try:
        host = urlparse(u).netloc or ""
        host = host.lower()
        host = host.replace("www.", "")
        return host
    except Exception:
        return ""


def _strip_google_redirect(u: str) -> str:
    return u


def _google_news_query(dominio: str) -> str:
    dominio = (dominio or "").strip()
    if not dominio:
        return ""
    return f"site:{dominio}"


def _google_news_rss_for_query(q: str) -> str:
    q = (q or "").strip()
    if not q:
        return ""
    return f"https://news.google.com/rss/search?q={q}&hl=pt-BR&gl=BR&ceid=BR:pt-419"


def _passa_filtro_politica(url: str, titulo: str, prefix: str) -> bool:
    if prefix and prefix.strip().upper() == "POLITICA":
        return True

    u = (url or "").lower()
    t = (titulo or "")
    if POLITICA_RE.search(u):
        return True
    if TITLE_KEYWORDS_RE.search(t):
        return True
    return False


@dataclass
class InputRow:
    dominio: str
    estado: str
    prefix: str
    row_idx: int


def _parse_input_sheet(df: pd.DataFrame) -> list[InputRow]:
    if df is None or df.empty:
        return []

    cols = [c.strip() for c in df.columns]
    df.columns = cols

    if COL_LINK not in df.columns or COL_ESTADO not in df.columns:
        raise ValueError(f"Planilha precisa ter colunas '{COL_LINK}' e '{COL_ESTADO}'.")

    inputs = []
    for i, r in df.iterrows():
        url = (r.get(COL_LINK) or "").strip()
        estado = _normalize_estado(r.get(COL_ESTADO) or "")
        if not url or not estado:
            continue

        dom = _domain_from_url(url)
        if not dom:
            continue

        prefix = str(r.get("Prefixo") or "").strip()
        if estado not in ESTADOS_SET:
            continue

        inputs.append(InputRow(dominio=dom, estado=estado, prefix=prefix, row_idx=i))

    return inputs


def _entry_datetime(entry) -> datetime | None:
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


def _hoje_local() -> date:
    return _now_local().date()


def _eh_ultimos_dias(dt_obj: datetime, dias: int) -> bool:
    if not dt_obj:
        return False
    dias = int(dias or 0)
    if dias <= 0:
        return True
    limite = _hoje_local() - timedelta(days=dias - 1)
    return dt_obj.date() >= limite


def _eh_hoje(dt_obj: datetime) -> bool:
    return _eh_ultimos_dias(dt_obj, 1)


def _gsheets_client_from_env():
    info_str = os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
    if info_str:
        info = json.loads(info_str)
    else:
        cred_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
        if not cred_path:
            cred_path = os.getenv("CRED_PATH", "").strip()
        if not cred_path:
            raise ValueError("Faltou credencial: GOOGLE_CREDENTIALS_JSON ou GOOGLE_APPLICATION_CREDENTIALS/CRED_PATH.")
        with open(cred_path, "r", encoding="utf-8") as f:
            info = json.load(f)

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


def _safe_gspread_call(fn, *args, **kwargs):
    tries = 6
    backoff = 1.5
    last = None
    for i in range(tries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last = e
            if _is_retryable_apierror(e):
                sleep_s = backoff ** i
                time.sleep(sleep_s)
                continue
            raise
    raise last


def _get_or_create_ws(sh, title: str, rows: int = 500, cols: int = 30):
    try:
        return sh.worksheet(title)
    except Exception:
        return _safe_gspread_call(sh.add_worksheet, title=title, rows=rows, cols=cols)


def _ensure_header(ws, header: list[str]):
    vals = _safe_gspread_call(ws.get_all_values)
    if not vals:
        _safe_gspread_call(ws.update, "A1", [header])
        return

    first = vals[0] if vals else []
    if [c.strip() for c in first] != header:
        _safe_gspread_call(ws.insert_row, header, 1)


def _read_input_sheet(gc, spreadsheet_id: str, aba: str) -> pd.DataFrame:
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(aba)
    data = ws.get_all_records()
    return pd.DataFrame(data)


def _dedup_existing_urls(ws) -> set[str]:
    vals = _safe_gspread_call(ws.get_all_values)
    if not vals or len(vals) < 2:
        return set()
    header = vals[0]
    try:
        idx = header.index("url")
    except Exception:
        return set()
    out = set()
    for row in vals[1:]:
        if idx < len(row):
            u = _normalize_url(row[idx])
            if u:
                out.add(u)
    return out


def _append_rows(ws, rows: list[dict], header: list[str], batch_rows: int = 60):
    if not rows:
        return

    values = []
    for r in rows:
        values.append([r.get(c, "") for c in header])

    for i in range(0, len(values), batch_rows):
        chunk = values[i : i + batch_rows]
        _safe_gspread_call(ws.append_rows, chunk, value_input_option="USER_ENTERED")


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
    else:
        ws_all = None
        existing_urls_all = set()

    for estado, sub in df.groupby("estado"):
        ws = _get_or_create_ws(sh, estado, rows=1200, cols=30)
        _ensure_header(ws, OUT_COLS)

        existing_urls_estado = _dedup_existing_urls(ws)

        rows_to_write = []
        for _, r in sub.iterrows():
            u = _normalize_url(r.get("url", ""))
            if not u:
                continue
            if u in existing_urls_estado:
                continue
            if ws_all is not None and u in existing_urls_all:
                continue

            rows_to_write.append(
                {
                    "data_publicacao": r.get("data_publicacao", ""),
                    "estado": r.get("estado", ""),
                    "dominio": r.get("dominio", ""),
                    "titulo": r.get("titulo", ""),
                    "url": r.get("url", ""),
                    "data_coleta": r.get("data_coleta", ""),
                }
            )

        if rows_to_write:
            _append_rows(ws, rows_to_write, OUT_COLS, batch_rows=BATCH_ROWS)
            print(f"{estado}: +{len(rows_to_write)} linhas")

            if ws_all is not None:
                _append_rows(ws_all, rows_to_write, OUT_COLS, batch_rows=BATCH_ROWS)
                for rr in rows_to_write:
                    uu = _normalize_url(rr.get("url", ""))
                    if uu:
                        existing_urls_all.add(uu)

        if PAUSA_ENTRE_ABAS:
            time.sleep(PAUSA_ENTRE_ABAS)


def coletar_por_uf():
    if not SPREADSHEET_ID:
        raise ValueError("SPREADSHEET_ID/PLANILHA_SUBNACIONAL não configurado.")

    gc = _gsheets_client_from_env()
    df_in = _read_input_sheet(gc, SPREADSHEET_ID, ABA_ENTRADA)
    inputs = _parse_input_sheet(df_in)

    if not inputs:
        print("Nada para coletar. Verifique a aba de entrada.")
        return

    by_state = {}
    for it in inputs:
        by_state.setdefault(it.estado, []).append(it)

    estados = sorted(by_state.keys())
    if MAX_ESTADOS_POR_RUN and MAX_ESTADOS_POR_RUN > 0:
        estados = estados[:MAX_ESTADOS_POR_RUN]

    start = time.time()
    budget_s = max(60, MAX_MINUTES_BUDGET * 60)

    rows = []
    seen_urls = set()

    for estado in estados:
        itens = by_state.get(estado, [])
        print(f"\nEstado: {estado} | itens: {len(itens)}")

        for it in itens:
            if (time.time() - start) > budget_s:
                print("Budget de tempo atingido. Encerrando cedo.")
                break

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

                dias_filtro = 1 if APENAS_HOJE else DIAS_RETROATIVOS
                if not _eh_ultimos_dias(dt_pub, dias_filtro):
                    continue

                if not _passa_filtro_politica(url, titulo, it.prefix):
                    continue

                seen_urls.add(url)

                rows.append(
                    {
                        "data_publicacao": dt_pub.strftime("%Y-%m-%d %H:%M:%S") if dt_pub else "",
                        "estado": estado,
                        "dominio": it.dominio,
                        "titulo": titulo,
                        "url": _truncate(url),
                        "data_coleta": _now_local().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )

            if PAUSA_ENTRE_ITENS:
                time.sleep(PAUSA_ENTRE_ITENS)

        if (time.time() - start) > budget_s:
            break

    df_out = pd.DataFrame(rows, columns=OUT_COLS)
    if df_out.empty:
        print("Sem notícias após filtros.")
        return

    gravar_por_estado(SPREADSHEET_ID, df_out)


if __name__ == "__main__":
    coletar_por_uf()
