"""
Joomla SEO Scanner - Pre-launch audit tool
Checks: title tags, meta descriptions, internal 404 links, noindex tags, robots.txt, XML sitemaps
Output: seo_report.xlsx
https://github.com/web-eau-net/joomla-seo-scanner
"""

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs, urlencode
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import xml.etree.ElementTree as ET
import sys
import time
from datetime import datetime

# --------------------------------------------
# CONFIGURATION ? edit this section
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
OUTPUT_FILE  = "seo_report.xlsx"

# Thresholds ? adjust to your needs
TITLE_MIN, TITLE_MAX = 40, 60
DESC_MIN,  DESC_MAX  = 120, 160

# Delay between requests in seconds ? increase if the server rate-limits you
DELAY = 0.5

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# --------------------------------------------
# EXCEL COLORS
# --------------------------------------------

C_RED_BG    = "FFDDDD"
C_ORANGE_BG = "FFF3CD"
C_GREEN_BG  = "D4EDDA"
C_BLUE_BG   = "D0E8FF"
C_HEADER_BG = "1F3864"
C_WHITE     = "FFFFFF"
C_RED_TXT   = "C0392B"
C_ORANGE_TXT= "856404"
C_GREEN_TXT = "155724"

# --------------------------------------------
# NETWORK HELPERS
# --------------------------------------------

session = requests.Session()
session.headers.update(HEADERS)

def get_page(url, timeout=15):
    try:
        return session.get(url, timeout=timeout, allow_redirects=True)
    except requests.RequestException:
        return None

def is_internal(url):
    parsed = urlparse(url)
    return parsed.netloc == "" or parsed.netloc == BASE_DOMAIN

def normalize_url(base, href):
    url = urljoin(base, href)
    parsed = urlparse(url)
    return parsed._replace(fragment="").geturl()

def canonical_url(url):
    """Canonical form for deduplication: lowercase, no trailing slash,
    strips pure Joomla session/language params (lang, tmpl, format, Itemid)."""
    parsed = urlparse(url)
    path = parsed.path.rstrip("/") or "/"
    params = parse_qs(parsed.query, keep_blank_values=True)
    for p in ["lang", "tmpl", "format", "Itemid"]:
        params.pop(p, None)
    query = urlencode(sorted(params.items())) if params else ""
    return parsed._replace(path=path, query=query, fragment="").geturl().lower()

# --------------------------------------------
# CRAWL
# --------------------------------------------

def crawl_site(start_urls):
    visited  = {}
    to_visit = set(start_urls)
    crawled  = set()

    print("\n" + "="*60)
    print("  CRAWLING...")
    print("="*60)

    while to_visit:
        url = to_visit.pop()
        canon = canonical_url(url)
        if canon in crawled:
            continue
        crawled.add(canon)

        print("  -> " + url)
        resp = get_page(url)
        time.sleep(DELAY)

        if resp is None:
            visited[canon] = {"status": None, "error": "Connection failed"}
            continue

        final_canon = canonical_url(resp.url)
        if final_canon in visited:
            continue

        visited[final_canon] = {
            "status":      resp.status_code,
            "title":       "",
            "description": "",
            "noindex":     False,
            "links":       [],
            "display_url": resp.url,
        }

        if resp.status_code != 200:
            continue

        soup = BeautifulSoup(resp.text, "html.parser")

        # Title tag
        title_tag = soup.find("title")
        visited[final_canon]["title"] = title_tag.get_text(strip=True) if title_tag else ""

        # Meta description
        desc_tag = soup.find("meta", attrs={"name": lambda v: v and v.lower() == "description"})
        visited[final_canon]["description"] = desc_tag.get("content", "").strip() if desc_tag else ""

        # Noindex
        robots_tag = soup.find("meta", attrs={"name": lambda v: v and v.lower() == "robots"})
        if robots_tag:
            visited[final_canon]["noindex"] = "noindex" in robots_tag.get("content", "").lower()

        # Internal links
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if href.startswith(("#", "mailto:", "tel:")):
                continue
            full = normalize_url(resp.url, href)
            c = canonical_url(full)
            if is_internal(full) and c not in crawled:
                to_visit.add(full)
            visited[final_canon]["links"].append(full)

    print("\n  [OK] " + str(len(visited)) + " pages crawled\n")
    return visited

# --------------------------------------------
# ANALYSIS
# --------------------------------------------

def analyse(pages):
    issues = []
    titles       = defaultdict(list)
    descriptions = defaultdict(list)

    for url, data in pages.items():
        if data.get("error") or data.get("status") != 200:
            continue
        t_key = " ".join(data["title"].lower().split())
        d_key = " ".join(data["description"].lower().split())
        if t_key:
            titles[t_key].append(url)
        if d_key:
            descriptions[d_key].append(url)
        data["title_key"] = t_key
        data["desc_key"]  = d_key

    def add(url, category, issue, fix, severity="[WARNING]"):
        issues.append({
            "url": url, "categorie": category,
            "erreur": issue, "correctif": fix, "gravite": severity,
        })

    for url, data in pages.items():

        # Connection error
        if data.get("error"):
            add(url, "Connection", "Page unreachable (network error)",
                "Check the URL and server", "[CRITICAL]")
            continue

        # HTTP error
        if data["status"] and data["status"] >= 400:
            add(url, "Dead link", "HTTP " + str(data["status"]) + " - Page not found",
                "Fix or remove this internal link", "[CRITICAL]")
            continue

        # Redirect
        if data["status"] and data["status"] >= 300:
            add(url, "Redirect", "HTTP " + str(data["status"]) + " - Redirect detected",
                "Update links to point directly to the final URL", "[WARNING]")

        # Title
        t = data.get("title", "")
        if not t:
            add(url, "Title tag", "Title tag missing",
                "Add a unique descriptive title (" + str(TITLE_MIN) + "-" + str(TITLE_MAX) + " chars)",
                "[CRITICAL]")
        else:
            if len(t) < TITLE_MIN:
                add(url, "Title tag",
                    "Title too short (" + str(len(t)) + " chars): " + t[:80],
                    "Expand the title to at least " + str(TITLE_MIN) + " characters",
                    "[WARNING]")
            elif len(t) > TITLE_MAX:
                add(url, "Title tag",
                    "Title too long (" + str(len(t)) + " chars): " + t[:80],
                    "Reduce the title to " + str(TITLE_MAX) + " characters maximum",
                    "[WARNING]")
            t_key = data.get("title_key", " ".join(t.lower().split()))
            if len(titles[t_key]) > 1 and titles[t_key][0] == url:
                urls_str = "\n".join("  - " + u for u in titles[t_key])
                add(url, "Title tag",
                    "Duplicate title on " + str(len(titles[t_key])) + " pages: " + t[:60] + "\n" + urls_str,
                    "Write a unique title for each listed page", "[CRITICAL]")

        # Meta description
        d = data.get("description", "")
        if not d:
            add(url, "Meta description", "Meta description missing",
                "Add a unique meta description (" + str(DESC_MIN) + "-" + str(DESC_MAX) + " chars)",
                "[WARNING]")
        else:
            if len(d) < DESC_MIN:
                add(url, "Meta description",
                    "Description too short (" + str(len(d)) + " chars): " + d[:80],
                    "Expand the description to at least " + str(DESC_MIN) + " characters",
                    "[WARNING]")
            elif len(d) > DESC_MAX:
                add(url, "Meta description",
                    "Description too long (" + str(len(d)) + " chars): " + d[:80],
                    "Reduce the description to " + str(DESC_MAX) + " characters maximum",
                    "[WARNING]")
            d_key = data.get("desc_key", " ".join(d.lower().split()))
            if len(descriptions[d_key]) > 1 and descriptions[d_key][0] == url:
                urls_str = "\n".join("  - " + u for u in descriptions[d_key])
                add(url, "Meta description",
                    "Duplicate description on " + str(len(descriptions[d_key])) + " pages: " + d[:60] + "\n" + urls_str,
                    "Write a unique meta description for each listed page", "[CRITICAL]")

        # Noindex
        if data.get("noindex"):
            add(url, "Noindex", "Noindex tag detected - page excluded from indexing",
                "Remove the noindex tag if this page should be indexed in production",
                "[CRITICAL]")

    # Internal 404 links ? grouped by dead URL, listing all source pages
    # Only checks URLs containing index.php (Joomla navigation links)
    dead_link_sources = defaultdict(list)
    for url, data in pages.items():
        if data.get("error") or data.get("status") != 200:
            continue
        for link in data.get("links", []):
            if not is_internal(link) or "index.php" not in link:
                continue
            lc = canonical_url(link)
            if lc in pages and pages[lc].get("status") == 404:
                if url not in dead_link_sources[lc]:
                    dead_link_sources[lc].append(url)

    for dead_url, source_pages in sorted(dead_link_sources.items()):
        sources_str = "\n".join("  - " + p for p in source_pages)
        add(dead_url, "Internal 404",
            "Dead link (" + str(len(source_pages)) + " source page(s)):\n" + sources_str,
            "Fix or remove this link on all listed source pages",
            "[CRITICAL]")

    return issues

# --------------------------------------------
# ROBOTS.TXT
# --------------------------------------------

def check_robots():
    issues = []
    resp = get_page(ROBOTS_URL)
    if resp is None or resp.status_code != 200:
        code = resp.status_code if resp else "N/A"
        issues.append({"url": ROBOTS_URL, "categorie": "robots.txt",
            "erreur": "robots.txt unreachable (HTTP " + str(code) + ")",
            "correctif": "Create a robots.txt file at the site root",
            "gravite": "[CRITICAL]"})
        return issues

    lines = resp.text.splitlines()
    has_ua  = any(l.strip().lower().startswith("user-agent") for l in lines)
    has_map = any("sitemap" in l.lower() for l in lines)
    disallow_all = any(
        l.strip().lower().startswith("disallow") and l.split(":", 1)[-1].strip() == "/"
        for l in lines
    )

    if not has_ua:
        issues.append({"url": ROBOTS_URL, "categorie": "robots.txt",
            "erreur": "No User-agent directive found",
            "correctif": "Add at minimum: User-agent: *", "gravite": "[WARNING]"})
    if not has_map:
        issues.append({"url": ROBOTS_URL, "categorie": "robots.txt",
            "erreur": "Sitemap reference missing",
            "correctif": "Add: Sitemap: https://yoursite.com/sitemap.xml",
            "gravite": "[WARNING]"})
    if disallow_all:
        issues.append({"url": ROBOTS_URL, "categorie": "robots.txt",
            "erreur": "Disallow: / detected - entire site blocked from crawlers",
            "correctif": "Remove or update the Disallow directive before going live",
            "gravite": "[CRITICAL]"})
    if not issues:
        issues.append({"url": ROBOTS_URL, "categorie": "robots.txt",
            "erreur": "[OK] No issues detected", "correctif": "", "gravite": "[OK]"})
    return issues

# --------------------------------------------
# SITEMAPS
# --------------------------------------------

def check_sitemaps():
    issues = []
    for sm_url in SITEMAP_URLS:
        resp = get_page(sm_url)
        if resp is None or resp.status_code != 200:
            code = resp.status_code if resp else "N/A"
            issues.append({"url": sm_url, "categorie": "sitemap.xml",
                "erreur": "Sitemap unreachable (HTTP " + str(code) + ")",
                "correctif": "Check sitemap generation and accessibility",
                "gravite": "[CRITICAL]"})
            continue
        try:
            root = ET.fromstring(resp.content)
        except ET.ParseError:
            issues.append({"url": sm_url, "categorie": "sitemap.xml",
                "erreur": "Invalid sitemap - malformed XML",
                "correctif": "Fix the XML structure of the sitemap",
                "gravite": "[CRITICAL]"})
            continue

        ns   = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}
        locs = root.findall(".//sm:loc", ns)

        if not locs:
            issues.append({"url": sm_url, "categorie": "sitemap.xml",
                "erreur": "Empty sitemap - no URLs found",
                "correctif": "Make sure the sitemap lists your site pages",
                "gravite": "[WARNING]"})
        else:
            sample = [loc.text.strip() for loc in locs[:5] if loc.text]
            dead = [u for u in sample if (r := get_page(u)) and r.status_code >= 400]
            if dead:
                for d in dead:
                    issues.append({"url": sm_url, "categorie": "sitemap.xml",
                        "erreur": "Sitemap URL returning error: " + d,
                        "correctif": "Remove dead URLs from the sitemap",
                        "gravite": "[CRITICAL]"})
            else:
                issues.append({"url": sm_url, "categorie": "sitemap.xml",
                    "erreur": "[OK] Valid sitemap - " + str(len(locs)) + " URL(s) listed",
                    "correctif": "", "gravite": "[OK]"})
    return issues

# --------------------------------------------
# EXCEL EXPORT
# --------------------------------------------

def export_excel(all_issues, pages):
    wb = openpyxl.Workbook()

    ws_sum = wb.active
    ws_sum.title = "Summary"

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    def hcell(ws, row, col, val, bg=C_HEADER_BG, fg=C_WHITE):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(bold=True, color=fg, name="Arial", size=11)
        c.fill      = PatternFill("solid", fgColor=bg)
        c.alignment = align_center
        return c

    def dcell(ws, row, col, val, bg=C_WHITE, fg="000000"):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(name="Arial", size=10, color=fg)
        c.fill      = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        return c

    thin   = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def bord(ws, row, col):
        ws.cell(row=row, column=col).border = border

    # Summary title
    ws_sum.merge_cells("A1:D1")
    t = ws_sum.cell(row=1, column=1,
        value="SEO Report - " + datetime.now().strftime("%Y-%m-%d %H:%M"))
    t.font      = Font(bold=True, name="Arial", size=14, color=C_WHITE)
    t.fill      = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = align_center
    ws_sum.row_dimensions[1].height = 30

    # Dynamic counters via COUNTIF
    NB = len(all_issues) + 1
    SH = "'Issues detected'"
    stats = [
        ("[!] Critical errors",
         '=COUNTIF(' + SH + '!A2:A' + str(NB) + ',"*CRITICAL*")',
         C_RED_BG, C_RED_TXT),
        ("[!] Warnings",
         '=COUNTIF(' + SH + '!A2:A' + str(NB) + ',"*WARNING*")',
         C_ORANGE_BG, C_ORANGE_TXT),
        ("[OK] Checks passed",
         '=COUNTIF(' + SH + '!A2:A' + str(NB) + ',"[OK]")',
         C_GREEN_BG, C_GREEN_TXT),
        ("[i] Pages crawled", len(pages), C_BLUE_BG, "1F3864"),
    ]

    for r, (label, val, bg, fg) in enumerate(stats, start=3):
        ws_sum.cell(row=r, column=1, value=label).font = Font(name="Arial", size=11, bold=True)
        ws_sum.cell(row=r, column=1).alignment = align_left
        ws_sum.cell(row=r, column=1).fill = PatternFill("solid", fgColor=bg)
        ws_sum.cell(row=r, column=2, value=val).font = Font(name="Arial", size=11, bold=True, color=fg)
        ws_sum.cell(row=r, column=2).fill = PatternFill("solid", fgColor=bg)
        ws_sum.cell(row=r, column=2).alignment = align_center

    ws_sum.column_dimensions["A"].width = 30
    ws_sum.column_dimensions["B"].width = 12

    # Issues worksheet
    ws = wb.create_sheet("Issues detected")
    headers = ["Severity", "Category", "Page URL", "Issue detected", "Fix"]
    col_w   = [16, 20, 55, 60, 55]
    for ci, (h, w) in enumerate(zip(headers, col_w), start=1):
        hcell(ws, 1, ci, h)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:E" + str(len(all_issues) + 1)

    GSTYLE = {
        "[CRITICAL]": (C_RED_BG,    C_RED_TXT),
        "[WARNING]":  (C_ORANGE_BG, C_ORANGE_TXT),
        "[OK]":       (C_GREEN_BG,  C_GREEN_TXT),
    }

    for ri, issue in enumerate(all_issues, start=2):
        grav = issue["gravite"]
        bg, fg = GSTYLE.get(grav, (C_WHITE, "000000"))
        ws.row_dimensions[ri].height = 32
        for ci, key in enumerate(["gravite", "categorie", "url", "erreur", "correctif"], start=1):
            dcell(ws, ri, ci, issue.get(key, ""),
                  bg=bg if ci == 1 else C_WHITE,
                  fg=fg if ci == 1 else "000000")
            bord(ws, ri, ci)
        url_val = issue.get("url", "")
        if url_val.startswith("http"):
            c = ws.cell(row=ri, column=3)
            c.hyperlink = url_val
            c.font = Font(name="Arial", size=10, color="0563C1", underline="single")

    # Crawled pages worksheet
    ws_p = wb.create_sheet("Crawled pages")
    ph = ["URL", "HTTP", "Title", "Title len", "Description", "Desc len", "Noindex"]
    pw = [55, 8, 50, 10, 60, 10, 10]
    for ci, (h, w) in enumerate(zip(ph, pw), start=1):
        hcell(ws_p, 1, ci, h)
        ws_p.column_dimensions[get_column_letter(ci)].width = w
    ws_p.freeze_panes = "A2"

    for ri, (url, data) in enumerate(pages.items(), start=2):
        ws_p.row_dimensions[ri].height = 20
        t  = data.get("title", "")
        d  = data.get("description", "")
        st = data.get("status", "")
        ni = "YES" if data.get("noindex") else ""
        for ci, val in enumerate([data.get("display_url", url), st, t,
                                   len(t) if t else 0, d, len(d) if d else 0, ni], start=1):
            dcell(ws_p, ri, ci, val)
            bord(ws_p, ri, ci)
        hc = ws_p.cell(row=ri, column=2)
        if st == 200:
            hc.fill = PatternFill("solid", fgColor=C_GREEN_BG)
        elif st and st >= 400:
            hc.fill = PatternFill("solid", fgColor=C_RED_BG)

    wb.save(OUTPUT_FILE)
    print("\n  [OK] Report exported: " + OUTPUT_FILE + "\n")

# --------------------------------------------
# MAIN
# --------------------------------------------

def main():
    print("\n" + "="*60)
    print("  JOOMLA SEO SCANNER - PRE-LAUNCH AUDIT")
    print("="*60)

    pages  = crawl_site(BASE_URLS)

    print("  RUNNING SEO ANALYSIS...")
    issues = analyse(pages)

    print("  Checking robots.txt...")
    issues += check_robots()

    print("  Checking sitemaps...")
    issues += check_sitemaps()

    order = {"[CRITICAL]": 0, "[WARNING]": 1, "[OK]": 2}
    issues.sort(key=lambda x: order.get(x["gravite"], 9))

    print("\n  GENERATING EXCEL REPORT...")
    export_excel(issues, pages)

    critiques = sum(1 for i in issues if i["gravite"] == "[CRITICAL]")
    avert     = sum(1 for i in issues if i["gravite"] == "[WARNING]")

    print("="*60)
    print("  Pages crawled:   " + str(len(pages)))
    print("  Critical errors: " + str(critiques))
    print("  Warnings:        " + str(avert))
    print("  Report:          " + OUTPUT_FILE)
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
