"""
SEO Re-analysis - Reads the existing report and regenerates with new thresholds.
No re-crawl needed - reads from the "Crawled pages" worksheet of a previous scan.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
from datetime import datetime
import sys
import os

# --------------------------------------------
# CONFIGURATION
# --------------------------------------------

INPUT_FILE  = "rapport_seo.xlsx"
OUTPUT_FILE = "rapport_seo_v2.xlsx"

# Thresholds
TITLE_MIN, TITLE_MAX = 50, 60
DESC_MIN,  DESC_MAX  = 120, 160

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
# READ EXISTING REPORT
# --------------------------------------------

def load_pages(filepath):
    print("  Reading " + filepath + "...")
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # Find the pages worksheet
    sheet_name = None
    for name in wb.sheetnames:
        if "crawl" in name.lower() or "pages" in name.lower():
            sheet_name = name
            break
    if not sheet_name:
        print("  ERROR: worksheet 'Pages crawlees' not found in " + filepath)
        print("  Available sheets: " + str(wb.sheetnames))
        sys.exit(1)

    ws = wb[sheet_name]
    print("  Sheet found: " + sheet_name)

    # Read headers
    headers = [str(ws.cell(row=1, column=c).value or "").strip().lower()
               for c in range(1, ws.max_column + 1)]
    print("  Columns: " + str(headers))

    # Map columns
    def col(name):
        for i, h in enumerate(headers):
            if name in h:
                return i + 1
        return None

    c_url    = col("url")
    c_http   = col("http")
    c_title  = col("title")
    c_lt     = col("long")  # Long.title
    c_desc   = col("desc")
    c_noind  = col("noindex")

    if not c_url:
        print("  ERROR: URL column not found")
        sys.exit(1)

    pages = {}
    for row in range(2, ws.max_row + 1):
        url    = str(ws.cell(row=row, column=c_url).value or "").strip()
        if not url or url == "None":
            continue
        status = ws.cell(row=row, column=c_http).value if c_http else None
        title  = str(ws.cell(row=row, column=c_title).value or "").strip() if c_title else ""
        desc   = str(ws.cell(row=row, column=c_desc).value or "").strip() if c_desc else ""
        noind  = str(ws.cell(row=row, column=c_noind).value or "").strip() if c_noind else ""

        # Clean "None" values
        if title == "None": title = ""
        if desc  == "None": desc  = ""

        try:
            status = int(status) if status else None
        except (ValueError, TypeError):
            status = None

        pages[url] = {
            "status":   status,
            "title":    title,
            "description": desc,
            "noindex":  noind.upper() == "OUI",
        }

    print("  " + str(len(pages)) + " pages lues\n")
    return pages

# --------------------------------------------
# ANALYSIS WITH NEW THRESHOLDS
# --------------------------------------------

def analyse(pages):
    issues = []
    titles       = defaultdict(list)
    descriptions = defaultdict(list)

    for url, data in pages.items():
        if data.get("status") != 200:
            continue
        t_key = " ".join(data["title"].lower().split())
        d_key = " ".join(data["description"].lower().split())
        if t_key:
            titles[t_key].append(url)
        if d_key:
            descriptions[d_key].append(url)
        data["title_key"] = t_key
        data["desc_key"]  = d_key

    def add(url, categorie, erreur, correctif, gravite="[WARNING]"):
        issues.append({
            "url": url, "categorie": categorie,
            "erreur": erreur, "correctif": correctif, "gravite": gravite,
        })

    for url, data in pages.items():

        if data.get("status") and data["status"] >= 400:
            add(url, "Dead link",
                "HTTP " + str(data["status"]) + " - Page not found",
                "Fix or remove this link", "[CRITICAL]")
            continue

        # Title
        t = data.get("title", "")
        if not t:
            add(url, "Title tag", "Title tag missing",
                "Add a unique descriptive title (" + str(TITLE_MIN) + "-" + str(TITLE_MAX) + " chars)",
                "[CRITICAL]")
        else:
            if len(t) < TITLE_MIN:
                add(url, "Title tag",
                    "Title too short (" + str(len(t)) + " car.) : " + t[:80],
                    "Expand the title to reach at least " + str(TITLE_MIN) + " characters",
                    "[WARNING]")
            elif len(t) > TITLE_MAX:
                add(url, "Title tag",
                    "Title too long (" + str(len(t)) + " car.) : " + t[:80],
                    "Reduce the title to " + str(TITLE_MAX) + " characters maximum",
                    "[WARNING]")
            t_key = data.get("title_key", " ".join(t.lower().split()))
            if len(titles[t_key]) > 1 and titles[t_key][0] == url:
                urls_str = "\n".join("  - " + u for u in titles[t_key])
                add(url, "Title tag",
                    "Duplicate title on " + str(len(titles[t_key])) + " pages: " + t[:60] + "\n" + urls_str,
                    "Rediger un title unique pour chaque page listee", "[CRITICAL]")

        # Description
        d = data.get("description", "")
        if not d:
            add(url, "Meta description", "Meta description missing",
                "Add a unique meta description (" + str(DESC_MIN) + "-" + str(DESC_MAX) + " chars)",
                "[WARNING]")
        else:
            if len(d) < DESC_MIN:
                add(url, "Meta description",
                    "Description too short (" + str(len(d)) + " car.) : " + d[:80],
                    "Expand the description to reach at least " + str(DESC_MIN) + " characters",
                    "[WARNING]")
            elif len(d) > DESC_MAX:
                add(url, "Meta description",
                    "Description too long (" + str(len(d)) + " car.) : " + d[:80],
                    "Reduce the description to " + str(DESC_MAX) + " characters maximum",
                    "[WARNING]")
            d_key = data.get("desc_key", " ".join(d.lower().split()))
            if len(descriptions[d_key]) > 1 and descriptions[d_key][0] == url:
                urls_str = "\n".join("  - " + u for u in descriptions[d_key])
                add(url, "Meta description",
                    "Duplicate description on " + str(len(descriptions[d_key])) + " pages: " + d[:60] + "\n" + urls_str,
                    "Rediger une meta description unique pour chaque page listee", "[CRITICAL]")

        # Noindex
        if data.get("noindex"):
            add(url, "Noindex",
                "Noindex tag detected - page excluded from indexing",
                "Remove the noindex tag if the page should be indexed in production",
                "[CRITICAL]")

    return issues

# --------------------------------------------
# EXCEL EXPORT
# --------------------------------------------

def export_excel(all_issues, pages, input_file):
    wb = openpyxl.Workbook()

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin   = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

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

    def bord(ws, row, col):
        ws.cell(row=row, column=col).border = border

    # ?? Onglet Resume ??
    ws_sum = wb.active
    ws_sum.title = "Summary"

    ws_sum.merge_cells("A1:D1")
    t = ws_sum.cell(row=1, column=1,
        value="SEO Report v2 - " + datetime.now().strftime("%d/%m/%Y %H:%M")
              + "  (thresholds: title " + str(TITLE_MIN) + "-" + str(TITLE_MAX)
              + " chars / desc " + str(DESC_MIN) + "-" + str(DESC_MAX) + " chars)")
    t.font      = Font(bold=True, name="Arial", size=12, color=C_WHITE)
    t.fill      = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = align_center
    ws_sum.row_dimensions[1].height = 30

    NB = len(all_issues) + 1
    SH = "'Issues detected'"
    stats = [
        ("[!] Critical errors",
         '=COUNTIF(' + SH + '!A2:A' + str(NB) + ',"*CRITIQUE*")',
         C_RED_BG, C_RED_TXT),
        ("[!] Warnings",
         '=COUNTIF(' + SH + '!A2:A' + str(NB) + ',"*AVERT*")',
         C_ORANGE_BG, C_ORANGE_TXT),
        ("[OK] Checks passed",
         '=COUNTIF(' + SH + '!A2:A' + str(NB) + ',"[OK]")',
         C_GREEN_BG, C_GREEN_TXT),
        ("[i] Pages analysed", len(pages), C_BLUE_BG, "1F3864"),
        ("Title min threshold (chars)",  TITLE_MIN, C_BLUE_BG, "1F3864"),
        ("Title max threshold (chars)",  TITLE_MAX, C_BLUE_BG, "1F3864"),
        ("Desc min threshold (chars)",   DESC_MIN,  C_BLUE_BG, "1F3864"),
        ("Desc max threshold (chars)",   DESC_MAX,  C_BLUE_BG, "1F3864"),
    ]

    for r, (label, val, bg, fg) in enumerate(stats, start=3):
        ws_sum.cell(row=r, column=1, value=label).font = Font(name="Arial", size=11, bold=True)
        ws_sum.cell(row=r, column=1).alignment = align_left
        ws_sum.cell(row=r, column=1).fill = PatternFill("solid", fgColor=bg)
        ws_sum.cell(row=r, column=2, value=val).font = Font(name="Arial", size=11, bold=True, color=fg)
        ws_sum.cell(row=r, column=2).fill = PatternFill("solid", fgColor=bg)
        ws_sum.cell(row=r, column=2).alignment = align_center

    ws_sum.column_dimensions["A"].width = 32
    ws_sum.column_dimensions["B"].width = 14

    # ?? Onglet Problemes ??
    ws = wb.create_sheet("Issues detected")
    headers = ["Severity", "Category", "Page URL", "Issue detected", "Fix"]
    col_w   = [16, 20, 55, 70, 55]
    for ci, (h, w) in enumerate(zip(headers, col_w), start=1):
        hcell(ws, 1, ci, h)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:E" + str(len(all_issues) + 1)

    GSTYLE = {
        "[CRITICAL]": (C_RED_BG,    C_RED_TXT),
        "[WARNING]":    (C_ORANGE_BG, C_ORANGE_TXT),
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

    # ?? Onglet Pages ??
    ws_p = wb.create_sheet("Crawled pages")
    ph = ["URL", "HTTP", "Title", "Title len", "Description", "Desc len", "Noindex"]
    pw = [55, 8, 50, 10, 60, 10, 10]
    for ci, (h, w) in enumerate(zip(ph, pw), start=1):
        hcell(ws_p, 1, ci, h)
        ws_p.column_dimensions[get_column_letter(ci)].width = w
    ws_p.freeze_panes = "A2"

    for ri, (url, data) in enumerate(pages.items(), start=2):
        t   = data.get("title", "")
        d   = data.get("description", "")
        st  = data.get("status", "")
        ni  = "OUI" if data.get("noindex") else ""
        lt  = len(t) if t else 0
        ld  = len(d) if d else 0
        for ci, val in enumerate([url, st, t, lt, d, ld, ni], start=1):
            dcell(ws_p, ri, ci, val)
            bord(ws_p, ri, ci)
        ws_p.row_dimensions[ri].height = 20
        hc = ws_p.cell(row=ri, column=2)
        if st == 200:
            hc.fill = PatternFill("solid", fgColor=C_GREEN_BG)
        elif st and st >= 400:
            hc.fill = PatternFill("solid", fgColor=C_RED_BG)
        # Colorer les longueurs hors seuils
        lt_c = ws_p.cell(row=ri, column=4)
        if t and (lt < TITLE_MIN or lt > TITLE_MAX):
            lt_c.fill = PatternFill("solid", fgColor=C_ORANGE_BG)
        ld_c = ws_p.cell(row=ri, column=6)
        if d and (ld < DESC_MIN or ld > DESC_MAX):
            ld_c.fill = PatternFill("solid", fgColor=C_ORANGE_BG)

    wb.save(OUTPUT_FILE)
    print("\n  [OK] Rapport exporte : " + OUTPUT_FILE + "\n")

# --------------------------------------------
# MAIN
# --------------------------------------------

def main():
    print("\n" + "="*60)
    print("  SEO RE-ANALYSIS - NEW THRESHOLDS")
    print("  Title: " + str(TITLE_MIN) + "-" + str(TITLE_MAX) + " car.")
    print("  Desc:  " + str(DESC_MIN)  + "-" + str(DESC_MAX)  + " car.")
    print("="*60)

    if not os.path.exists(INPUT_FILE):
        print("  ERROR: file '" + INPUT_FILE + "' not found in current folder.")
        print("  Make sure rapport_seo.xlsx is in the same folder as this script.")
        sys.exit(1)

    pages  = load_pages(INPUT_FILE)

    print("  RUNNING ANALYSIS...")
    issues = analyse(pages)

    order = {"[CRITICAL]": 0, "[WARNING]": 1, "[OK]": 2}
    issues.sort(key=lambda x: order.get(x["gravite"], 9))

    print("  GENERATING REPORT...")
    export_excel(issues, pages, INPUT_FILE)

    critiques = sum(1 for i in issues if i["gravite"] == "[CRITICAL]")
    avert     = sum(1 for i in issues if i["gravite"] == "[WARNING]")

    print("="*60)
    print("  Pages analysed:  " + str(len(pages)))
    print("  Critical errors: " + str(critiques))
    print("  Warnings:        " + str(avert))
    print("  Report:          " + OUTPUT_FILE)
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
