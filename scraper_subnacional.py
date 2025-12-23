import os
import re
import time
import unicodedata
from datetime import datetime, timedelta, date
from urllib.parse import urlparse, quote_plus

import requests
import feedparser
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials


def _norm_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def _strip_accents(s: str) -> str:
    s = s or ""
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))


def _now_str(tz_offset_hours: int) -> str:
    return (datetime.utcnow() + timedelta(hours=tz_offset_hours)).strftime("%Y-%m-%d %H:%M:%S")


def _parse_dt(entry) -> datetime | None:
    if getattr(entry, "published_parsed", None):
        try:
            return datetime(*entry.published_parsed[:6])
        except Exception:
            return None
    if getattr(entry, "updated_parsed", None):
        try:
            return datetime(*entry.updated_parsed[:6])
        except Exception:
            return None
    for k in ("published", "updated"):
        if k in entry:
            try:
                dt = pd.to_datetime(entry.get(k), errors="coerce")
                if pd.isna(dt):
                    return None
                return dt.to_pydatetime()
            except Exception:
                return None
    return None


def _is_feed_url(u: str) -> bool:
    u = (u or "").lower().strip()
    return any(x in u for x in ("/feed", "format=feed", "rss", ".xml", "atom"))


def _has_path_beyond_root(u: str) -> bool:
    try:
        p = urlparse(u)
        path = (p.path or "").strip("/")
        return bool(path)
    except Exception:
        return False


def _google_news_rss_url(query: str) -> str:
    q = quote_plus(query)
    return f"https://news.google.com/rss/search?q={q}&hl=pt-BR&gl=BR&ceid=BR:pt-419"


def _domain_and_path(u: str) -> tuple[str, str]:
    p = urlparse(u.strip())
    domain = (p.netloc or "").lower()
    path = (p.path or "").strip()
    if not path.startswith("/"):
        path = "/" + path if path else ""
    return domain, path


def _candidate_feeds_from_section_url(section_url: str) -> list[str]:
    base = section_url.strip().rstrip("/")
    candidates = [
        base + "/feed/",
        base + "/feed",
        base + "/rss",
        base + "/rss/",
        base + "/rss.xml",
        base + "?feed=rss",
        base + "?format=feed&type=rss",
        base + "?output=rss",
        base + "?type=rss",
        base + "/index.xml",
        base + "/feed.xml",
    ]
    out, seen = [], set()
    for c in candidates:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _looks_like_rss(text: str) -> bool:
    t = (text or "").lower()
    return ("<rss" in t) or ("<feed" in t) or ("atom" in t and "<feed" in t)


def _try_fetch_rss(url: str, timeout: int = 15) -> str | None:
    try:
        r = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code != 200:
            return None
        ct = (r.headers.get("content-type") or "").lower()
        if ("xml" in ct) or ("rss" in ct) or ("atom" in ct) or _looks_like_rss(r.text[:5000]):
            return url
        return None
    except Exception:
        return None


def _ensure_worksheet(sh, title: str, columns: list[str], header_row: int) -> gspread.Worksheet:
    safe_title = title[:31]
    try:
        ws = sh.worksheet(safe_title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=safe_title, rows=200, cols=max(10, len(columns)))

    values = ws.get_all_values()
    need_min_rows = max(header_row, 3)
    if ws.row_count < need_min_rows:
        ws.resize(rows=need_min_rows)

    # garante que o header exista exatamente na linha header_row (deixa as linhas acima livres pro "botão")
    header_a1 = f"A{header_row}"
    header_range_end_col = len(columns)

    # checa se já tem header nessa linha
    row_vals = []
    try:
        row_vals = ws.row_values(header_row)
    except Exception:
        row_vals = []

    if not row_vals or not any((c or "").strip() for c in row_vals):
        ws.resize(rows=max(ws.row_count, header_row + 1), cols=max(ws.col_count, header_range_end_col))
        ws.update(header_a1, [columns])

    return ws


def _get_existing_urls(ws: gspread.Worksheet, header_row: int) -> set[str]:
    values = ws.get_all_values()
    if not values or len(values) < header_row + 1:
        return set()

    header = [h.strip() for h in (values[header_row - 1] if len(values) >= header_row else [])]
    try:
        idx = [h.lower() for h in header].index("url")
    except ValueError:
        return set()

    urls = set()
    data_start = header_row  # 1-based row index; in list it's header_row
    for row in values[data_start:]:
        if idx < len(row):
            u = (row[idx] or "").strip()
            if u:
                urls.add(u)
    return urls


def _insert_rows_for_top_write(sh, ws: gspread.Worksheet, n: int, header_row: int):
    if n <= 0:
        return

    # queremos inserir as novas linhas imediatamente após o header
    start_index_0b = header_row  # 0-based: header_row (1-based) vira header_row-1; inserir após header => startIndex = header_row
    # exemplo: header_row=3 => header está em index 2; após header => startIndex=3
    if ws.row_count < (header_row + 1):
        ws.resize(rows=header_row + 1)

    sheet_id = ws._properties["sheetId"]
    req = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index_0b,
                        "endIndex": start_index_0b + n,
                    },
                    "inheritFromBefore": True,
                }
            }
        ]
    }
    sh.batch_update(req)


def _values_update(ws: gspread.Worksheet, start_row: int, start_col: int, values_2d: list[list[str]]):
    if not values_2d:
        return
    n_rows = len(values_2d)
    n_cols = max(len(r) for r in values_2d)

    def col_letter(n: int) -> str:
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    start_a1 = f"{col_letter(start_col)}{start_row}"
    end_a1 = f"{col_letter(start_col + n_cols - 1)}{start_row + n_rows - 1}"
    ws.update(f"{start_a1}:{end_a1}", values_2d, value_input_option="RAW")


def coletar_por_uf(
    spreadsheet_id: str,
    aba_entrada: str = "deduplicado",
    col_link: str = "Link",
    col_estado: str = "Estado",
    tz_offset_hours: int = -3,
    apenas_hoje: bool = True,
    dias: int = 2,
    max_por_feed: int = 100,
    pausa_seg: float = 0.15,
    sleep_ufs: float = 0.5,
    sleep_writes: float = 1.2,
    header_row: int = 3,  # linha 1 e 2 livres pro botão/espacinho; header na 3
):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json").strip()
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)

    sh = gc.open_by_key(spreadsheet_id)

    ws_in = sh.worksheet(aba_entrada)
    raw = ws_in.get_all_values()
    if not raw or len(raw) < 2:
        raise ValueError(f"A aba '{aba_entrada}' não trouxe dados (verifique header e conteúdo).")

    header = [h.strip() for h in raw[0]]
    rows = raw[1:]
    df = pd.DataFrame(rows, columns=header)

    cols_norm = {c.strip().lower(): c for c in df.columns}
    if col_link.lower() not in cols_norm or col_estado.lower() not in cols_norm:
        raise ValueError(f"Esperado '{col_link}' e '{col_estado}' na aba '{aba_entrada}'.")

    link_col = cols_norm[col_link.lower()]
    estado_col = cols_norm[col_estado.lower()]

    df = df[[link_col, estado_col]].copy()
    df[link_col] = df[link_col].astype(str).map(lambda x: x.strip())
    df[estado_col] = df[estado_col].astype(str).map(lambda x: x.strip())

    df = df[(df[link_col] != "") & (df[estado_col] != "")]
    if df.empty:
        raise ValueError(f"A aba '{aba_entrada}' está sem linhas válidas em '{col_link}'/'{col_estado}'.")

    df = df.drop_duplicates(subset=[link_col, estado_col], keep="first")

    out_cols = ["data_coleta", "estado", "fonte", "titulo", "url", "data_publicacao", "origem_link"]

    hoje = date.today()
    dt_min = datetime.utcnow() - timedelta(days=dias)

    def ok_dt(dt_obj: datetime | None) -> bool:
        if not dt_obj:
            return True
        local_dt = dt_obj + timedelta(hours=tz_offset_hours)
        if apenas_hoje:
            return local_dt.date() == hoje
        return dt_obj >= dt_min

    seen_global = set()

    for estado, g in df.groupby(estado_col, sort=True):
        estado_tab = _norm_spaces(estado).upper()

        links = g[link_col].tolist()
        links = [l for l in links if l and l.lower().startswith("http")]
        if not links:
            continue

        print(f"\nEstado: {estado_tab} | links: {len(links)}")

        noticias = []
        for origem_link in links:
            origem_link = origem_link.strip()
            domain, path = _domain_and_path(origem_link)

            feed_urls = []
            if _is_feed_url(origem_link):
                feed_urls = [origem_link]
            else:
                if _has_path_beyond_root(origem_link):
                    candidates = _candidate_feeds_from_section_url(origem_link)
                    for c in candidates:
                        ok = _try_fetch_rss(c)
                        if ok:
                            feed_urls.append(ok)
                            break

                if not feed_urls:
                    if domain:
                        if path and path != "/":
                            q = f"site:{domain}{path}"
                        else:
                            q = f"(politica OR poder OR eleicao OR eleições OR eleitoral) site:{domain}"
                        feed_urls = [_google_news_rss_url(q)]

            for fu in feed_urls:
                try:
                    fp = feedparser.parse(fu)
                    entries = fp.entries[:max_por_feed]
                    print(f"  {domain or origem_link} -> {len(entries)} entradas")

                    for e in entries:
                        title = _norm_spaces(e.get("title", ""))
                        url = (e.get("link", "") or "").strip()
                        if not url or not title:
                            continue

                        if not _has_path_beyond_root(origem_link):
                            t_low = _strip_accents(title).lower()
                            if not re.search(
                                r"\b(politic|poder|elei(c|ç)ao|elei(c|ç)oes|governador|prefeito|candid|partid|camara|assembleia|senad|deputad)\b",
                                t_low,
                            ):
                                continue

                        if url in seen_global:
                            continue

                        dt_pub = _parse_dt(e)
                        if dt_pub and not ok_dt(dt_pub):
                            continue

                        noticias.append(
                            {
                                "data_coleta": _now_str(tz_offset_hours),
                                "estado": estado_tab,
                                "fonte": domain or origem_link,
                                "titulo": title,
                                "url": url,
                                "data_publicacao": (
                                    (dt_pub + timedelta(hours=tz_offset_hours)).strftime("%Y-%m-%d %H:%M:%S")
                                    if dt_pub
                                    else ""
                                ),
                                "origem_link": origem_link,
                            }
                        )
                        seen_global.add(url)

                    if pausa_seg:
                        time.sleep(pausa_seg)

                except Exception as ex:
                    print(f"  erro lendo feed: {fu} | {ex}")

        if not noticias:
            print(f"✓ {estado_tab}: nada novo (coleta vazia).")
            time.sleep(sleep_ufs)
            continue

        ws = _ensure_worksheet(sh, estado_tab, out_cols, header_row=header_row)
        existing_urls = _get_existing_urls(ws, header_row=header_row)

        df_out = pd.DataFrame(noticias)
        df_out = df_out.drop_duplicates(subset=["url"], keep="first")
        df_out = df_out[~df_out["url"].astype(str).isin(existing_urls)]

        if df_out.empty:
            print(f"✓ {estado_tab}: nada novo para inserir (dedup).")
            time.sleep(sleep_ufs)
            continue

        if "data_publicacao" in df_out.columns:
            df_out = df_out.sort_values(by=["data_publicacao"], ascending=False, kind="stable")

        values_2d = df_out[out_cols].astype(str).values.tolist()

        try:
            # escreve sempre abaixo do header_row, deixando as linhas acima (1-2) pro botão/espacinho
            _insert_rows_for_top_write(sh, ws, len(values_2d), header_row=header_row)
            _values_update(ws, start_row=header_row + 1, start_col=1, values_2d=values_2d)
            print(f"✓ {estado_tab}: inseridas {len(values_2d)} linhas no topo.")
        except Exception as ex:
            print(f"✗ {estado_tab}: erro ao escrever ({ex}).")
            raise

        time.sleep(sleep_writes)
        time.sleep(sleep_ufs)


if __name__ == "__main__":
    SPREADSHEET_ID = os.getenv("PLANILHA_SUBNACIONAL", "").strip()
    if not SPREADSHEET_ID:
        raise SystemExit("Defina SPREADSHEET_ID no ambiente (key da planilha).")

    ABA_ENTRADA = os.getenv("ABA_ENTRADA", "deduplicado").strip()
    TZ_OFFSET_HOURS = int(os.getenv("TZ_OFFSET_HOURS", "-3").strip())
    APENAS_HOJE = os.getenv("APENAS_HOJE", "1").strip().lower() in ("1", "true", "yes", "on")
    DIAS = int(os.getenv("DIAS", "2").strip())
    MAX_POR_FEED = int(os.getenv("MAX_POR_FEED", "100").strip())
    PAUSA_SEG = float(os.getenv("PAUSA_SEG", "0.15").strip())
    SLEEP_UFS = float(os.getenv("SLEEP_UFS", "0.5").strip())
    SLEEP_WRITES = float(os.getenv("SLEEP_WRITES", "1.2").strip())
    HEADER_ROW = int(os.getenv("HEADER_ROW", "3").strip())

    coletar_por_uf(
        spreadsheet_id=SPREADSHEET_ID,
        aba_entrada=ABA_ENTRADA,
        tz_offset_hours=TZ_OFFSET_HOURS,
        apenas_hoje=APENAS_HOJE,
        dias=DIAS,
        max_por_feed=MAX_POR_FEED,
        pausa_seg=PAUSA_SEG,
        sleep_ufs=SLEEP_UFS,
        sleep_writes=SLEEP_WRITES,
        header_row=HEADER_ROW,
    )
