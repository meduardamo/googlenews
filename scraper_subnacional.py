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
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials


MAX_CELL_CHARS = int(os.getenv("MAX_CELL_CHARS", "47000"))
TZ_OFFSET_HOURS = int(os.getenv("TZ_OFFSET_HOURS", "-3"))

PLANILHA_ENV = os.getenv("PLANILHA_SUBNACIONAL", "").strip() or os.getenv("SPREADSHEET_ID", "").strip()
ABA_ENTRADA = os.getenv("ABA_ENTRADA", "deduplicado").strip()
COL_LINK = os.getenv("COL_LINK", "Link").strip()
COL_ESTADO = os.getenv("COL_ESTADO", "Estado").strip()

APENAS_HOJE = os.getenv("APENAS_HOJE", "1").strip().lower() in ("1", "true", "yes", "on")
MAX_PER_DOMAIN = int(os.getenv("MAX_PER_DOMAIN", "100"))
PAUSA_ENTRE_DOMINIOS = float(os.getenv("PAUSA_ENTRE_DOMINIOS", "0.2"))

WRITE_CONSOLIDADO = os.getenv("WRITE_CONSOLIDADO", "0").strip().lower() in ("1", "true", "yes", "on")
ABA_CONSOLIDADO = os.getenv("ABA_CONSOLIDADO", "links_com_estado_monitoramento").strip()

BATCH_INSERT_SIZE = int(os.getenv("BATCH_INSERT_SIZE", "50"))
PAUSA_BETWEEN_BATCH_WRITES = float(os.getenv("PAUSA_BETWEEN_BATCH_WRITES", "2.0"))
PAUSA_BETWEEN_WS = float(os.getenv("PAUSA_BETWEEN_WS", "1.0"))
MAX_WRITES_PER_RUN = int(os.getenv("MAX_WRITES_PER_RUN", "1000000"))

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


def _read_ws_df(ws):
    values = ws.get_all_values()
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
    values = ws.get_all_values()
    if not values or not values[0] or not any(v.strip() for v in values[0]):
        tmp = pd.DataFrame(columns=columns)
        set_with_dataframe(ws, tmp, include_index=False, include_column_header=True, resize=True)
        return True
    return False


def _get_or_create_ws(sh, title, rows=200, cols=30):
    name = (title or "SEM_NOME").strip()[:31]
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=name, rows=rows, cols=cols)


def _ensure_min_rows(ws, min_rows: int):
    if ws.row_count < min_rows:
        ws.resize(rows=min_rows)


def _ensure_capacity_for_insert(ws, n_new_rows: int, insert_at_row: int = 2):
    _ensure_min_rows(ws, max(2, insert_at_row))
    ws.resize(rows=ws.row_count + n_new_rows)


def _chunk_df(df: pd.DataFrame, chunk_size: int):
    if df.empty:
        return []
    return [df.iloc[i : i + chunk_size].copy() for i in range(0, len(df), chunk_size)]


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

    for estado, dominios in by_state.items():
        dominios = sorted(set(dominios))
        print(f"\nEstado: {estado} | domínios: {len(dominios)}")

        for dom in dominios:
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


def _dedup_existing_urls(ws) -> set:
    existing = _read_ws_df(ws)
    if existing.empty:
        return set()

    url_col = next((c for c in existing.columns if c.strip().lower() == "url"), None)
    if not url_col:
        return set()

    return set(existing[url_col].astype(str).tolist())


def _safe_insert_rows_and_write(ws, df_to_write: pd.DataFrame, insert_at_row: int = 2):
    if df_to_write.empty:
        return 0

    writes_done = 0
    chunks = _chunk_df(df_to_write, BATCH_INSERT_SIZE)

    for i, chunk in enumerate(chunks, start=1):
        if writes_done >= MAX_WRITES_PER_RUN:
            print("Atingiu MAX_WRITES_PER_RUN, parando para evitar quota.")
            break

        _ensure_capacity_for_insert(ws, len(chunk), insert_at_row=insert_at_row)
        ws.insert_rows([[]] * len(chunk), row=insert_at_row, value_input_option="RAW")

        set_with_dataframe(
            ws,
            chunk,
            row=insert_at_row,
            col=1,
            include_index=False,
            include_column_header=False,
            resize=False,
        )

        writes_done += 2
        if PAUSA_BETWEEN_BATCH_WRITES:
            time.sleep(PAUSA_BETWEEN_BATCH_WRITES)

    return writes_done


def gravar_por_estado(spreadsheet_id: str, df: pd.DataFrame):
    if df.empty:
        print("Sem notícias para gravar.")
        return

    gc = _gsheets_client_from_env()
    sh = gc.open_by_key(spreadsheet_id)

    total_writes = 0

    if WRITE_CONSOLIDADO:
        ws_all = _get_or_create_ws(sh, ABA_CONSOLIDADO, rows=500, cols=30)
        _ensure_header(ws_all, OUT_COLS)

        existing_urls_all = _dedup_existing_urls(ws_all)
        df_all = df[~df["url"].astype(str).isin(existing_urls_all)].copy()

        if not df_all.empty:
            total_writes += _safe_insert_rows_and_write(ws_all, df_all, insert_at_row=2)
            print(f"✓ Consolidado: inseridas {len(df_all)} linhas em {ws_all.title}")
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

        total_writes += _safe_insert_rows_and_write(ws, sub2, insert_at_row=2)
        print(f"✓ {estado_tab}: inseridas {len(sub2)} linhas no topo.")

        if total_writes >= MAX_WRITES_PER_RUN:
            print("Atingiu MAX_WRITES_PER_RUN, encerrando execução.")
            break

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
        print("Defina PLANILHA_SUBNACIONAL no ambiente.", file=sys.stderr)
        sys.exit(1)

    coletar_por_uf(PLANILHA_ENV)


if __name__ == "__main__":
    main()
