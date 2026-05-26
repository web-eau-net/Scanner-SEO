# Joomla SEO Scanner

A Python-based SEO audit tool for Joomla websites. Runs locally on your machine and crawls your site — including development servers not publicly accessible — to generate a detailed Excel report.

Built for the [web-eau.net](https://web-eau.net) redesign project. Read the full story in the companion article: [Optimizing a Joomla site before launch](https://web-eau.net/en/blog/optimizing-joomla-site-before-launch).

---

## What it checks

| Check | Details |
|---|---|
| **Title tag** | Missing, duplicate, too short (< 40 chars) or too long (> 60 chars) |
| **Meta description** | Missing, duplicate, too short (< 120 chars) or too long (> 160 chars) |
| **Internal 404 links** | Dead links grouped by target URL, with all source pages listed |
| **Noindex tags** | Pages accidentally excluded from indexing |
| **robots.txt** | Presence, blocking `Disallow: /`, missing Sitemap reference |
| **XML Sitemaps** | XML validity, accessibility of listed URLs |

---

## Output

The script generates a `rapport_seo.xlsx` file with three worksheets:

- **Erreurs a traiter** — real issues only, color-coded by error type, with filters and clickable URLs
- **Tous les problemes** — full unfiltered report
- **Pages crawlees** — complete page inventory with HTTP status, title/description lengths (highlighted in orange when out of range)

![Excel report screenshot](docs/screenshot.png)

---

## Requirements

- Python 3.10 or higher
- Works on Windows, macOS and Linux

---

## Installation

**1. Clone the repository**
```bash
git clone https://github.com/web-eau-net/joomla-seo-scanner.git
cd joomla-seo-scanner
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

On some systems (Linux/macOS) you may need:
```bash
pip3 install -r requirements.txt
```

On Windows with multiple Python versions:
```bash
python -m pip install -r requirements.txt
```

---

## Configuration

Open `scanner_seo.py` and edit the configuration block at the top of the file:

```python
# --------------------------------------------
# CONFIGURATION
# --------------------------------------------

BASE_URLS = [
    "https://yoursite.com/",
    "https://yoursite.com/en/",   # add or remove language versions as needed
]

ROBOTS_URL   = "https://yoursite.com/robots.txt"
SITEMAP_URLS = [
    "https://yoursite.com/sitemap.xml",
    "https://yoursite.com/sitemap-en.xml",  # add or remove as needed
]

BASE_DOMAIN  = "yoursite.com"
OUTPUT_FILE  = "rapport_seo.xlsx"

# Thresholds — adjust to your needs
TITLE_MIN, TITLE_MAX = 40, 60
DESC_MIN,  DESC_MAX  = 120, 160

# Delay between requests (seconds) — increase if the server rate-limits you
DELAY = 0.5
```

---

## Usage

Place `scanner_seo.py` and `requirements.txt` in the same folder, then run:

```bash
python scanner_seo.py
```

The script will:
1. Crawl all internal pages starting from `BASE_URLS`
2. Analyse each page for SEO issues
3. Check `robots.txt` and XML sitemaps
4. Generate `rapport_seo.xlsx` in the same folder

**Expected runtime:** approximately 15–30 minutes for a 1,000–2,000 page site (depending on server speed and `DELAY` setting).

Progress is displayed in the terminal:

```
============================================================
  SCANNER SEO - AUDIT PRE-MISE EN LIGNE
============================================================

  CRAWL EN COURS...
============================================================
  -> https://yoursite.com/
  -> https://yoursite.com/about/
  -> https://yoursite.com/blog/
  ...
  [OK] 1471 pages analysees

  ANALYSE SEO EN COURS...
  Verification robots.txt...
  Verification sitemaps...

  GENERATION DU RAPPORT EXCEL...
  [OK] Rapport exporte : rapport_seo.xlsx

============================================================
  Pages crawlees   : 1471
  Erreurs critiques: 65
  Avertissements   : 42
  Rapport          : rapport_seo.xlsx
============================================================
```

---

## Handling false positives (Joomla-specific)

Joomla generates multiple URL variants for the same page (`?lang=fr`, `?Itemid=123`, `?start=15`, `?type=rss`...). The scanner handles this automatically via a `canonical_url()` function that normalises all variants before deduplication.

Pagination pages, RSS feeds and static files (`.pdf`, `.jpg`, `.png`) are automatically excluded from title/description duplicate checks.

If you still see unexpected duplicates, check whether your Joomla site uses custom URL parameters that should be added to the exclusion list in the `canonical_url()` function.

---

## Known limitations

- **JavaScript-rendered content** is not supported — the scanner reads raw HTML. Pages that load content via AJAX after the initial HTML response will not be fully analysed.
- **Authentication-protected pages** are not crawled. If your dev server requires HTTP Basic Auth, add credentials to the `session.headers` in the script.
- **External links** are not checked — only internal links pointing to `BASE_DOMAIN` are analysed for 404 errors.

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| `requests` | 2.32.3 | HTTP requests and crawling |
| `beautifulsoup4` | 4.12.3 | HTML parsing |
| `openpyxl` | 3.1.5 | Excel report generation |
| `lxml` | 5.2.2 | XML sitemap parsing |

---

## Compatibility

Tested on :
- Joomla 5.x and 6.x
- Python 3.10, 3.11, 3.12, 3.14
- Windows 10/11, macOS, Ubuntu 22.04+

Should work on any CMS or website — not Joomla-specific except for the URL normalisation logic.

---

## License

GPL v3 — see [LICENSE](LICENSE) for details.

---

## Related projects

- **[Errors MetaData](https://github.com/web-eau-net/Errors-MetaData)** — Joomla module to audit missing and incorrect metadata directly from the admin dashboard. Companion tool to this scanner.

---

## Contributing

Issues and pull requests are welcome. If you find a bug or have a feature request, please open an issue with as much detail as possible (Joomla version, Python version, error message).

---

*Built with ❤ in Brittany — [web-eau.net](https://web-eau.net)*
