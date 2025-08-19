h = it.find(["h3", "h4"])
a = h.find("a", href=True) if h else None
title = a.get_text(strip=True) if a else (h.get_text(strip=True) if h else None)

href = a["href"] if a else None
if not href:
    continue

# normaliza link relativo do GNews
base = "https://news.google.com"
if href.startswith("./"):
    href = base + href[1:]
elif href.startswith("/"):
    href = base + href

# descarte links que apontam para tópicos/edições do próprio Google News
parsed = urlparse(href)
if "news.google.com" in parsed.netloc and any(
    seg in parsed.path for seg in ("/topics", "/publications", "/headlines", "/stories")
):
    # são páginas de cluster; pulamos
    continue

# URL final real (via Selenium + fallback requests, como já está no seu código)
url_final = resolve_final_url_selenium(driver, href)
fp        = url_fingerprint(url_final)
dominio   = urlparse(url_final).netloc if url_final else ""

# se o título parece ser "genérico" do cluster ou ficou vazio, puxa o do publisher
needs_real_title = (
    not title
    or title.lower().startswith("notícias sobre")
    or (" • " in title and "notícias sobre" in title.lower())
    or "news.google.com" in urlparse(url_final).netloc
)
if needs_real_title:
    real = fetch_publisher_title(url_final)
    if real:
        title = real

# se ainda não tiver título, descarta (evita linha feia)
if not title:
    continue
