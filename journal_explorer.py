"""
Journal Explorer
================
A one-file desktop-style app (runs in your browser) that:

  1. Downloads every paper of the journals you pick from OpenAlex + Crossref (both free, no login).
  2. Lets you search / sort / filter them and see keyword statistics.
  3. Exports an Excel file laid out like a Web of Science export (same 67 column tags, one sheet
     per journal, plus a Summary sheet with keyword counts).

Libraries doing the heavy lifting (so this file stays small):
  * nicegui   - the web UI (tables, charts, buttons, progress bars) - no HTML/JS written by hand
  * pyalex    - OpenAlex client (paging, journal search, abstract rebuilding)
  * pandas    - the data table + Excel writing (via openpyxl)
  * requests  - talks to Crossref (only needed for the cited-reference lists)

Run:  python journal_explorer.py        (or double-click the built exe/app)
"""
import collections
import html
import json
import random
import re
import sys
import threading
import time
import unicodedata
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

import numpy as np
import pandas as pd
import requests
import xlsxwriter   # noqa: F401  (pandas loads it by name; the import makes the exe bundle it)
import reportlab    # noqa: F401  (used inside export_pdf; the import makes the exe bundle it)
from nicegui import run, ui
from pyalex import Sources, Works

# ----------------------------------------------------------------------------------------------
# 1. WHERE THINGS LIVE
# ----------------------------------------------------------------------------------------------
# When packaged into an exe, files must go NEXT TO the exe (not inside its temp folder).
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
DATA_DIR = APP_DIR / "data"          # one .pkl per journal + the settings file
DATA_DIR.mkdir(exist_ok=True)
CONFIG_FILE = DATA_DIR / "journals.json"
EXCEL_FILE = APP_DIR / "Journals.xlsx"

# The four journals the app ships with. "abbr" becomes the Excel sheet name.
DEFAULT_JOURNALS = [
    {"name": "Strategic Management Journal", "abbr": "SMJ", "oa_id": "S102949365", "issn": "0143-2095"},
    {"name": "Long Range Planning", "abbr": "LRP", "oa_id": "S9435936", "issn": "0024-6301"},
    {"name": "Strategic Organization", "abbr": "SO", "oa_id": "S11394615", "issn": "1476-1270"},
    {"name": "Strategy Science", "abbr": "StrSci", "oa_id": "S4210183760", "issn": "2333-2050"},
]

# The Web of Science column tags, in the same order as the original workbook.
WOS_COLUMNS = ("PT AU BA BE GP AF BF CA TI SO SE BS LA DT CT CY CL SP HO DE ID AB C1 RP EM RI OI FU FX "
               "CR NR TC Z9 U1 U2 PU PI PA SN EI BN J9 JI PD PY VL IS PN SU SI MA BP EP AR DI D2 EA PG "
               "WC SC GA UT PM OA HC HP DA").split()
# Columns we add on the right (not in WoS): journal, our type guess, the OpenAlex-vs-Crossref cross-check
# (Sources / Check / Conflicts), Crossref's own citation count, and the OpenAlex id.
EXTRA_COLUMNS = ["Journal", "Type", "Sources", "Check", "Conflicts", "CitedCrossref", "OpenAlexID"]

# Document types shown as filter chips. Everything else is grouped under "Other".
DOC_TYPES = ["Article", "Review", "Book Review", "Editorial Material", "Letter", "Correction", "Front/Back Matter", "Other"]
DEFAULT_ON = {"Article", "Review"}   # what the viewer shows at first; Excel still gets everything

# Title patterns that reveal a type OpenAlex mislabels as a plain "article".
TITLE_RULES = [
    (re.compile(r"^(book )?reviews?\b[:\s]|\bbook review\b|^review of\b", re.I), "Book Review"),
    (re.compile(r"\b(corrigendum|erratum|correction to|retraction)\b", re.I), "Correction"),
    (re.compile(r"\b(editorial board|front matter|back matter|table of contents|contents|index|cover|masthead|issue information)\b", re.I), "Front/Back Matter"),
    (re.compile(r"^(editor'?s? (note|comment|introduction)|editorial)\b|\bintroduction to the (special|forum)", re.I), "Editorial Material"),
]
OA_TYPE_MAP = {"article": "Article", "review": "Review", "editorial": "Editorial Material", "letter": "Letter",
               "erratum": "Correction", "retraction": "Correction", "paratext": "Front/Back Matter"}


def classify(oa_type: str, title: str, is_paratext: bool) -> str:
    """Best-guess document type: title patterns first (they catch mislabelled book reviews), then OpenAlex's own tag."""
    if is_paratext:
        return "Front/Back Matter"
    for pattern, label in TITLE_RULES:
        if pattern.search(title or ""):
            return label
    return OA_TYPE_MAP.get(oa_type or "", "Other")


# ----------------------------------------------------------------------------------------------
# 2. SETTINGS (which journals are tracked)
# ----------------------------------------------------------------------------------------------
def load_config() -> dict:
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    cfg = {"journals": DEFAULT_JOURNALS, "last_fetch": {}}
    save_config(cfg)
    return cfg


def save_config(cfg: dict) -> None:
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def pkl(j: dict) -> Path:
    return DATA_DIR / f"{j['abbr']}.pkl"


VOCAB: frozenset = frozenset()


def load_all() -> pd.DataFrame:
    """Stack every downloaded journal into one table (empty table if nothing downloaded yet)."""
    frames = [pd.read_pickle(pkl(j)) for j in CONFIG["journals"] if pkl(j).exists()]
    if not frames:
        return pd.DataFrame(columns=WOS_COLUMNS + EXTRA_COLUMNS)
    df = pd.concat(frames, ignore_index=True)
    df["_search"] = (df["TI"].fillna("") + " " + df["AB"].fillna("") + " " + df["DE"].fillna("")).str.lower()
    # the keywords of each paper as a set, so "does this paper have keyword X" is instant (Explore tab)
    df["_kw"] = df["DE"].fillna("").map(lambda s: frozenset(k.strip() for k in s.split(";") if k.strip()))
    # lower-case copies of the searchable fields, made once here so every filter click is fast
    df["_kwl"] = df["_kw"].map(lambda ks: frozenset(k.lower() for k in ks))
    df["_kwtext"] = df["DE"].fillna("").str.lower()
    df["_ti"] = df["TI"].fillna("").str.lower()
    df["_ab"] = df["AB"].fillna("").str.lower()
    df["_au"] = (df["AF"].fillna("") + " " + df["AU"].fillna("")).str.lower()
    global VOCAB                                                 # every keyword in the data (lower case)
    VOCAB = frozenset(k for ks in df["_kwl"] for k in ks)
    return df


# ----------------------------------------------------------------------------------------------
# 3. DOWNLOADING (OpenAlex for the papers, Crossref for the cited references)
# ----------------------------------------------------------------------------------------------
def short_name(full: str) -> str:
    """'Michael Bikard' -> 'Bikard, M' (WoS style)."""
    parts = full.replace(".", " ").split()
    return f"{parts[-1]}, {''.join(p[0] for p in parts[:-1])}" if len(parts) > 1 else full


def openalex_rows(j: dict, progress) -> list[dict]:
    """Page through all OpenAlex works of one journal and turn each into a WoS-style row."""
    fields = ["id", "doi", "title", "publication_year", "publication_date", "type", "is_paratext", "language",
              "cited_by_count", "authorships", "keywords", "abstract_inverted_index", "biblio", "open_access",
              "referenced_works_count", "primary_location", "funders"]
    query = Works().filter(primary_location={"source": {"id": j["oa_id"]}}).select(fields)
    total = query.count()
    rows = []
    for page in query.paginate(per_page=200, n_max=None):
        for w in page:
            rows.append(to_row(w, j))
        progress(j["abbr"], "oa", len(rows), total)
    return rows


def to_row(w: dict, j: dict) -> dict:
    """Map one OpenAlex work onto the WoS column tags. Fields OpenAlex lacks stay empty."""
    auth = w.get("authorships") or []
    names = [(a.get("author") or {}).get("display_name") or a.get("raw_author_name") or "" for a in auth]
    # C1 = affiliations, WoS style: "[Author] Institution, Country"
    c1 = "; ".join(f"[{n}] " + ", ".join(filter(None, [i.get("display_name"), i.get("country_code")]))
                   for n, a in zip(names, auth) for i in (a.get("institutions") or [])[:1])
    orcids = "; ".join(f"{n}/{(a.get('author') or {}).get('orcid', '').split('/')[-1]}"
                       for n, a in zip(names, auth) if (a.get("author") or {}).get("orcid"))
    bib = w.get("biblio") or {}
    title = w.get("title") or ""
    doi = (w.get("doi") or "").replace("https://doi.org/", "")
    row = dict.fromkeys(WOS_COLUMNS)
    row.update({
        "PT": "J", "AU": "; ".join(short_name(n) for n in names), "AF": "; ".join(names),
        "TI": title, "SO": j["name"].upper(), "LA": "English" if w.get("language") == "en" else w.get("language"),
        "DT": classify(w.get("type"), title, bool(w.get("is_paratext"))),
        "DE": "; ".join(k["display_name"] for k in (w.get("keywords") or [])),
        "AB": w["abstract"],                       # pyalex rebuilds the text from the inverted index
        "C1": c1, "OI": orcids, "FU": "; ".join(f["display_name"] for f in (w.get("funders") or [])),
        "NR": w.get("referenced_works_count"), "TC": w.get("cited_by_count"), "Z9": w.get("cited_by_count"),
        "SN": j["issn"], "PD": w.get("publication_date"), "PY": w.get("publication_year"),
        "VL": bib.get("volume"), "IS": bib.get("issue"), "BP": bib.get("first_page"), "EP": bib.get("last_page"),
        "DI": doi, "OA": (w.get("open_access") or {}).get("oa_status"), "UT": w["id"].split("/")[-1],
        "Journal": j["abbr"], "OpenAlexID": w["id"].split("/")[-1],
    })
    row["Type"] = row["DT"]
    return row


def strip_tags(text: str | None) -> str | None:
    """Crossref titles/abstracts come as JATS/XML with HTML entities (&amp;); keep only the plain text."""
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", text or ""))).strip() or None


def crossref_records(j: dict, progress) -> dict[str, dict]:
    """Every Crossref record of one journal (by ISSN), keyed by lower-case DOI. Crossref only knows what the
    publisher deposited: that is why references / abstracts / volumes are sometimes missing."""
    out, cursor = {}, "*"
    while True:
        r = requests.get("https://api.crossref.org/works", timeout=120, params={
            "filter": f"issn:{j['issn']}", "rows": 500, "cursor": cursor,
            "select": "DOI,title,author,issued,volume,issue,page,type,reference,article-number,publisher,abstract,"
                      "is-referenced-by-count,references-count"})
        r.raise_for_status()
        msg = r.json()["message"]
        if not msg["items"]:
            return out
        for it in msg["items"]:
            refs = []
            for ref in it.get("reference", []):
                # Use Crossref's free-text citation when it exists, else stitch one from the parts.
                refs.append(ref.get("unstructured") or ", ".join(filter(None, [
                    ref.get("author"), ref.get("year"), ref.get("journal-title") or ref.get("volume-title"),
                    ref.get("volume") and "V" + ref["volume"], ref.get("first-page") and "P" + ref["first-page"],
                    ref.get("DOI") and "DOI " + ref["DOI"]])))
            pages = (it.get("page") or "").split("-")
            authors = [f"{a.get('family', '')}, {a.get('given', '')}".strip(", ") for a in it.get("author", [])]
            year = ((it.get("issued") or {}).get("date-parts") or [[None]])[0][0]
            out[it["DOI"].lower()] = {
                "DI": it["DOI"], "TI": strip_tags((it.get("title") or [""])[0]) or "", "PY": year, "VL": it.get("volume"),
                "IS": it.get("issue"), "BP": pages[0] or None, "EP": pages[1] if len(pages) > 1 else None,
                "AR": it.get("article-number"), "PU": it.get("publisher"), "AB": strip_tags(it.get("abstract")),
                "CR": "\n".join(refs)[:32000] or None, "TC": it.get("is-referenced-by-count"),
                "NR": it.get("references-count"), "AF": "; ".join(authors), "type": it.get("type"),
                "n_auth": len(authors)}
        cursor = msg["next-cursor"]
        progress(j["abbr"], "cr", len(out), msg["total-results"])


def norm(text) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def compare(row: dict, c: dict) -> list[tuple]:
    """Compare the fields BOTH sources have. Returns (field, openalex value, crossref value) for each that differs.
    A field is skipped when either side is empty (nothing to compare). Citation counts are deliberately NOT
    compared: OpenAlex always counts more citing sources than Crossref, so they would differ on every row."""
    diffs, compared = [], 0
    n_oa = len(row["AF"].split("; ")) if row.get("AF") else None
    pairs = [("year", row["PY"], c["PY"], lambda a, b: abs(int(a) - int(b)) <= 1),   # online-first vs print issue year
              ("volume", row["VL"], c["VL"], None), ("issue", row["IS"], c["IS"], None),
             ("first page", row["BP"], c["BP"], None), ("number of authors", n_oa, c["n_auth"] or None, None),
             # titles: agree if one contains the other after stripping punctuation (subtitle / casing differences)
             ("title", row["TI"], c["TI"], lambda a, b: norm(a) in norm(b) or norm(b) in norm(a))]
    for field, a, b, same in pairs:
        if a in (None, "", 0) or b in (None, "", 0):
            continue
        compared += 1
        if not (same(a, b) if same else str(a) == str(b)):
            diffs.append((field, a, b))
    row["_compared"] = compared
    return diffs


def merge(oa_rows: list[dict], cr: dict[str, dict], j: dict) -> pd.DataFrame:
    """Combine both sources into one table and mark how far they agree (the Check column)."""
    seen, out = set(), []
    for row in oa_rows:
        key = (row["DI"] or "").lower()
        c = cr.get(key)
        if c is None:
            row.update(Sources="OpenAlex only", Check="\u25fb OpenAlex only", Conflicts="")
        else:
            seen.add(key)
            diffs = compare(row, c)
            row["Sources"], row["CitedCrossref"] = "Both", c["TC"]
            if diffs:
                row["Check"] = "\u26a0\ufe0f Conflict: " + ", ".join(d[0] for d in diffs)
                row["Conflicts"] = " | ".join(f"{f}: OpenAlex {a} vs Crossref {b}" for f, a, b in diffs) + " (Crossref value kept)"
                # In a 250-paper check against scite, Crossref's year / volume / issue / first page were the right
                # ones in every case that could be settled (OpenAlex often holds the online-first year), so keep those.
                for f, tag in (("year", "PY"), ("volume", "VL"), ("issue", "IS"), ("first page", "BP")):
                    if any(d[0] == f for d in diffs):
                        row[tag] = c[tag]
            elif row["_compared"] > 0:
                row["Check"], row["Conflicts"] = "\u2705 Both agree", ""
            else:
                row["Check"], row["Conflicts"] = "\u25fb Both, nothing to compare", ""
            for tag in ("VL", "IS", "BP", "EP", "AR", "PU", "AB", "CR"):   # fill gaps from the other source
                row[tag] = row.get(tag) or c.get(tag)
        row.pop("_compared", None)
        out.append(row)
    for key, c in cr.items():                      # records only Crossref has
        if key in seen:
            continue
        row = dict.fromkeys(WOS_COLUMNS)
        row.update({k: c.get(k) for k in ("TI", "AF", "PY", "VL", "IS", "BP", "EP", "AR", "PU", "AB", "CR", "NR", "TC", "DI")})
        row.update(PT="J", SO=j["name"].upper(), SN=j["issn"], AU=c["AF"], Z9=c["TC"], Journal=j["abbr"],
                   DT=classify("article" if c["type"] == "journal-article" else "other", c["TI"], False),
                   Sources="Crossref only", Check="\u25fb Crossref only", Conflicts="", CitedCrossref=c["TC"])
        row["Type"] = row["DT"]
        out.append(row)
    return pd.DataFrame(out).reindex(columns=WOS_COLUMNS + EXTRA_COLUMNS)


def refresh_journal(j: dict, progress) -> None:
    """Download one journal from both sources AT THE SAME TIME (two threads) and save the merged table.
    Always a full re-pull: it takes minutes, and it keeps citation counts fresh."""
    with ThreadPoolExecutor(max_workers=2) as pool:
        openalex = pool.submit(openalex_rows, j, progress)
        crossref = pool.submit(crossref_records, j, progress)
        oa_rows, cr = openalex.result(), crossref.result()
    df = merge(oa_rows, cr, j)
    df.to_pickle(pkl(j))
    CONFIG["last_fetch"][j["abbr"]] = date.today().isoformat()
    save_config(CONFIG)


# ----------------------------------------------------------------------------------------------
# 4. EXCEL EXPORT (looks like the WoS workbook: plain Calibri, frozen header row, filter, 70% zoom)
# ----------------------------------------------------------------------------------------------
SUMMARY_KEYWORDS = ["Climate change", "Natural environment", "Environmental", "Green", "ESG", "Stakeholder",
                    "Shareholder value", "Performance", "Financial performance", "Agency theory", "Resilience",
                    "Growth", "Profitability", "De-growth", "decline", "competitive advantage"]


def col_letter(tag: str) -> str:
    """Excel column letter of a WoS tag (PT is A, AU is B ...)."""
    n, s = WOS_COLUMNS.index(tag) + 1, ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# Characters that are illegal in Excel's XML (control chars, lone surrogates, U+FFFE / U+FFFF).
ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
EXPORT_LOCK = threading.Lock()   # held while Journals.xlsx is being built
EXCEL_CELL_LIMIT = 32000          # Excel's hard limit is 32,767 characters per cell


def clean_cell(v):
    """Make one value safe for Excel: strip illegal characters, cap the length, and stop text that begins with
    '=' from being written as a formula. Any of these left in place makes Excel say the file is corrupt."""
    if not isinstance(v, str):
        return v
    v = ILLEGAL_XML.sub("", v)[:EXCEL_CELL_LIMIT]
    return " " + v if v.startswith("=") else v


def sheet_names() -> dict[str, str]:
    """Excel-safe, unique sheet name per journal abbreviation (no []:*?/\, max 31 chars, not the reserved 'History')."""
    used, out = {"summary"}, {}
    for j in CONFIG["journals"]:
        name = re.sub(r"[\[\]:*?/\\]", "", j["abbr"]).strip("'")[:31] or "Journal"
        if name.lower() in used or name.lower() == "history":
            name = name[:28] + "_" + str(len(used))
        used.add(name.lower())
        out[j["abbr"]] = name
    return out


def top_keywords(frames: list[pd.DataFrame], n: int = 25) -> list[str]:
    """The n most used keywords in the given journal tables (one count per paper)."""
    kw = pd.concat([f["DE"] for f in frames]).dropna()
    return kw.str.split("; ").explode().str.strip().value_counts().head(n).index.tolist()


def export_excel(selection: pd.DataFrame | None = None) -> Path:
    """Write Journals.xlsx - or, given a `selection` of papers (the Explore tab), the SAME workbook containing only those
    papers, saved as Explore_papers_<time>.xlsx with an extra 'Filters' sheet saying how they were chosen.
    It is built in a temporary file and only moved into place once EVERYTHING succeeded,
    so a failure part-way can never leave a half-written (corrupt) Journals.xlsx behind.
    xlsxwriter is used because it can store the *calculated* value next to every formula: the Summary then shows its
    numbers even where Excel does not recalculate (Protected View, previewers, Mac Quick Look)."""
    # Two exports writing the same file at once produce a damaged zip that Excel calls "corrupt" - so only one may run.
    if not EXPORT_LOCK.acquire(blocking=False):
        raise RuntimeError("An export is already running - please wait for it to finish.")
    target = EXCEL_FILE if selection is None else APP_DIR / f"Explore_papers_{datetime.now():%Y%m%d_%H%M}.xlsx"
    tmp = target.with_name(f"{target.stem}.building-{uuid.uuid4().hex[:8]}.xlsx")
    names = sheet_names()
    keep = None if selection is None else set(selection["OpenAlexID"].fillna(selection["DI"]))   # one id per paper
    try:
        # strings_to_formulas / strings_to_urls off: otherwise text starting "=" or looking like a link (every DOI URL!)
        # is turned into a formula / hyperlink, and Excel refuses files with too many hyperlinks.
        with pd.ExcelWriter(tmp, engine="xlsxwriter",
                            engine_kwargs={"options": {"strings_to_formulas": False, "strings_to_urls": False}}) as xl:
            book = xl.book
            summary = book.add_worksheet("Summary")                      # created first, so it is the first tab
            if selection is not None:                                    # how this selection was made
                info = book.add_worksheet("Filters")
                lines = describe_filters()
                info.write(0, 0, "Papers selected in Journal Explorer")
                info.write(1, 0, f"{len(selection):,} papers - created {datetime.now():%Y-%m-%d %H:%M}")
                for r, line in enumerate(lines, start=3):
                    info.write(r, 0, line)
                info.write(len(lines) + 4, 0, "Data: OpenAlex + Crossref - see 'About this data' in the app.")
                info.set_column(0, 0, 110)
            fills = {c: book.add_format({"bg_color": "#" + col}) for c, col in
                     {"\u2705": "C6EFCE", "\u26a0": "FFEB9C", "other": "EDEDED"}.items()}
            data = []                                                    # (journal, sheet name, table) for the Summary
            for j in CONFIG["journals"]:
                if not pkl(j).exists():
                    continue
                raw = pd.read_pickle(pkl(j))[WOS_COLUMNS + EXTRA_COLUMNS]
                if keep is not None:                                     # only the selected papers
                    raw = raw[raw["OpenAlexID"].fillna(raw["DI"]).isin(keep)]
                    if raw.empty:
                        continue
                df = raw.map(clean_cell)
                sh = names[j["abbr"]]
                df.to_excel(xl, sheet_name=sh, index=False)
                ws = xl.sheets[sh]
                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)
                ws.set_zoom(70)
                cc = df.columns.get_loc("Check")                         # colour the cross-check: green / amber / grey
                for r, v in enumerate(df["Check"].fillna(""), start=1):
                    ws.write_string(r, cc, v, fills.get(v[:1], fills["other"]))
                data.append((j, sh, df))

            # ---- Summary: keyword hits per journal in Title / Abstract / Keywords (live formulas + cached values) ----
            summary.write("C1", "JOURNALS")
            summary.write("B2", "Data: OpenAlex + Crossref, cross-checked (column 'Check' on each sheet). Not equivalent to "
                                "Scopus / Web of Science - see 'About this data' in the app.")
            summary.write("B4", "No of articles")
            summary.write("B5", "Keywords (the most used ones; type over any of them or use the empty rows)")
            keywords = top_keywords([df for _, _, df in data]) + [""] * 5  # last rows are free for your own keywords
            for r, kw in enumerate(keywords, start=6):
                summary.write(r - 1, 1, kw)
            for i, (j, sh, df) in enumerate(data):
                c0 = 2 + i * 4                                           # first column of this journal's block
                q = "'" + sh + "'"                                       # always quoted: names with spaces are legal
                summary.write(2, c0, j["name"])
                summary.write_formula(3, c0, f'=COUNTIF({q}!$A:$A,"=J")', None, len(df))
                for k, (label, tag) in enumerate([("Title", "TI"), ("Abstract", "AB"), ("Keywords", "DE")]):
                    summary.write(4, c0 + k, label)
                    text = df[tag].fillna("").str.lower()
                    letter = col_letter(tag)
                    for r, kw in enumerate(keywords, start=6):
                        # same rule as COUNTIF "*kw*": case-insensitive "contains"; the free rows show 0 until you type
                        hits = int(text.str.contains(kw.lower(), regex=False).sum()) if kw else 0
                        summary.write_formula(r - 1, c0 + k, f'=COUNTIF({q}!${letter}:${letter},"*"&Summary!$B{r}&"*")', None, hits)
            summary.set_column(1, 1, 38)
        tmp.replace(target)                                              # only now does the real file change
    finally:
        tmp.unlink(missing_ok=True)                                      # never leave the temp file behind
        EXPORT_LOCK.release()
    return target


# ----------------------------------------------------------------------------------------------
# 4b. THE "ABOUT THIS DATA" DISCLAIMER (shown on first run, always reachable from the header)
# ----------------------------------------------------------------------------------------------
DISCLAIMER = """
### Read this before you rely on the numbers

This tool uses two **free, automated** sources instead of a curated citation database such as **Scopus** or
**Web of Science**. That makes it free and easy to extend, but the data is less reliable.

**OpenAlex** (papers, authors, keywords, citation counts)
- Built by *automatic matching*, so it can contain wrong author or affiliation matches, duplicate or merged
  records, and missing papers. Nothing is hand-checked.
- **Keywords are machine-generated topics**, not the keywords the authors chose (Scopus and WoS carry author keywords).
  Expect broad or odd labels such as "Computer science" or "Resource (disambiguation)".
- **Abstracts are missing for many papers** (some publishers do not allow them to be shared).
- Document types are guesses: book reviews and front matter are often tagged as plain "articles". This app adds a
  title-based guess, which is also imperfect.
- **Citation counts are not comparable with Scopus / WoS.** Each database counts different citing sources.

**Crossref** (volume / issue / pages, cited references, its own citation count)
- Only contains what the **publisher chose to deposit**. Reference lists, abstracts and page numbers are often missing.
- No keywords, no reliable affiliations or author IDs, and no subject categories.
- Its citation count only counts citations from other Crossref members, so it is usually lower.

**Scopus / Web of Science** are curated: consistent author IDs and affiliations, author + indexed keywords,
near-complete reference lists, checked document types and subject categories, and a fixed coverage policy.
They are paid (or need a subscription) and have their own gaps and counting rules.

**The "Sources agree?" column**
- We download every paper from *both* sources, match them by DOI, and compare year (a 1-year gap is tolerated:
  online-first vs print issue), volume, issue, first page, title and number of authors.
  ✅ = both agree, ⚠️ = they conflict (the dialog for the paper shows exactly where), ◻ = only one source has it.
- Agreement is a *useful hint, not proof*: OpenAlex re-uses a lot of Crossref's metadata, so the two are **not
  independent**. Two matching wrong values are still wrong. Where only one source has the paper, treat it with more care.
- Citation counts are shown side by side and are never used for the agree / conflict flag.
- Missing fields in OpenAlex are filled from Crossref when Crossref has them.

**What differences to expect - a real check (Sept 2026)**
We compared 250 papers with scite (80 SMJ, 80 Long Range Planning, 25 Strategic Organization, 25 Strategy Science,
plus 40 that this app had flagged as conflicts). All 250 were found there.
- **Where both sources agreed (210 papers): very reliable.** Volume, issue and first page matched scite in all but
  one comparison; the year matched for 204 of 209 and was 1 year off for the other 5 (online-first vs print year).
  No real title disagreements: the only differences were papers with no title in any source (issue and front-matter
  records, and a few articles) and scite's own "&amp;" text.
- **Where the sources conflicted (40 papers), Crossref was right.** In 35 the year differed by 2 or more years,
  nearly all in *Long Range Planning*: OpenAlex gives the year a paper first appeared **online** (e.g. 2013) while
  Crossref and scite give the **print-issue year** (2015), which is what Scopus and Web of Science use. In every
  such case scite matched Crossref, so the app keeps Crossref's value and shows both in the paper's detail view.
  Two SMJ papers also had the wrong volume / issue in OpenAlex (23 vs 28). Author-count conflicts could not be
  settled and are left as OpenAlex's value.
- **Citation counts:** OpenAlex was within 25% of scite for 178 of 192 papers (median ratio 1.0). Crossref's count
  was about 0.84 of scite's on the median, and less than half of scite's for 18 of 192 papers, so treat it as a
  lower bound. Expect all three to differ from Scopus / Web of Science.
- **How complete is it? It depends on the era and the publisher** (articles and reviews in these four journals):
  papers up to 1995 have an abstract in only 12% of cases, a reference list in 38%, and an author list in 70%;
  papers from 2016 on have 81%, 98% and 99%. By publisher: Long Range Planning (Elsevier) has abstracts for just
  3% of its papers, while SMJ, Strategic Organization and Strategy Science have 93-95%. Issue-level and front-matter
  records often have no title at all.
- This is a 250-paper spot check on four journals, not a guarantee for other journals you add.

For anything that will be published or used for a decision, verify a sample against Scopus / WoS or the publisher.
"""

# ----------------------------------------------------------------------------------------------
# 5. THE UI
# ----------------------------------------------------------------------------------------------
CONFIG = load_config()
DF = load_all()                       # the table the viewer shows; reloaded after every download
N_PICKS = 10                          # how many keywords the slot machine shows at once
OPS = ["AND", "OR", "AND NOT"]        # how a term is joined to the ones before it
CHAIN_LABELS = {"keywords": "Keywords", "title": "Title", "authors": "Authors", "abstract": "Abstract"}


def year_bounds() -> tuple[int, int]:
    """Earliest and latest publication year in the data (or a sensible default when nothing is downloaded)."""
    y = pd.to_numeric(DF["PY"], errors="coerce").dropna() if not DF.empty else pd.Series(dtype=float)
    return (int(y.min()), int(y.max())) if len(y) else (1960, date.today().year)


YB = year_bounds()

# ONE shared filter state: the Papers tab, the Explore tab, the charts and the pop-up editors all read and write this,
# so they can never disagree. "chains" hold the AND / OR / AND NOT term lists (as long as you like) per text field.
STATE = {"types": set(DEFAULT_ON), "journals": set(), "check": "", "q": "", "years": {"min": YB[0], "max": YB[1]},
         "chains": {f: [] for f in CHAIN_LABELS},
         "busy": False, "error": "", "started": 0.0, "jobs": {}}
# State that only belongs to the Explore tab (the slot machine).
EX = {"picks": [], "dirty": True, "surprise": 0.6, "min_papers": 2, "next_op": "AND", "chart_mode": "count", "kw_mode": "most", "pdf_abs": True}


def term_mask(d: pd.DataFrame, field: str, text: str) -> pd.Series:
    """Which papers contain one term. Keywords: an exact keyword when the term is a known one (so 'Process' does not
    also match 'Process management'), otherwise a partial match. Title / authors / abstract: partial match."""
    t = text.strip().lower()
    if field == "keywords":
        if t in VOCAB:
            return d["_kwl"].map(lambda s: t in s)
        return d["_kwtext"].str.contains(t, regex=False)
    return d[{"title": "_ti", "authors": "_au", "abstract": "_ab"}[field]].str.contains(t, regex=False)


def chain_mask(d: pd.DataFrame, field: str, chain: list[dict]):
    """Evaluate 'a AND b OR c AND NOT d' (AND binds tighter than OR, like everywhere else). None when the chain is empty."""
    chain = [c for c in chain if c["text"].strip()]
    if not chain or d.empty:
        return None
    groups, cur = [], None
    for i, c in enumerate(chain):
        m = term_mask(d, field, c["text"]).to_numpy()
        if c["op"] == "AND NOT":
            m = ~m
        if i == 0 or c["op"] == "OR":                            # an OR starts a new group
            if cur is not None:
                groups.append(cur)
            cur = m
        else:
            cur = cur & m
    groups.append(cur)
    out = groups[0]
    for g in groups[1:]:
        out = out | g
    return out


def filtered(skip: tuple = ()) -> pd.DataFrame:
    """Apply the shared filters. `skip` leaves out chains (the charts need 'everything except the keyword chain')."""
    d = DF
    if d.empty:
        return d
    d = d[d["Type"].isin(STATE["types"])]
    if STATE["journals"]:
        d = d[d["Journal"].isin(STATE["journals"])]
    if STATE["check"]:                                           # confidence filter (see Check column)
        d = d[d["Check"].fillna("").str.startswith(STATE["check"])]
    if (STATE["years"]["min"], STATE["years"]["max"]) != YB:     # only when narrowed, so papers without a year stay in
        year = pd.to_numeric(d["PY"], errors="coerce")
        d = d[(year >= STATE["years"]["min"]) & (year <= STATE["years"]["max"])]
    for term in STATE["q"].lower().split():
        d = d[d["_search"].str.contains(term, regex=False)]
    for field, chain in STATE["chains"].items():
        if field not in skip:
            m = chain_mask(d, field, chain)
            if m is not None:
                d = d[m]
    return d


# ---- PDF export of the papers left after the filters and picks (Explore tab) ----
def describe_filters() -> list[str]:
    """The active filters as readable lines - printed on the first page of the PDF so it explains itself."""
    lines = []
    for field, chain in STATE["chains"].items():
        chain = [c for c in chain if c["text"].strip()]
        if chain:
            txt = "".join((("NOT " if c["op"] == "AND NOT" else "") if i == 0 else f" {c['op']} ") + c["text"] for i, c in enumerate(chain))
            lines.append(f"{CHAIN_LABELS[field]}: {txt}")
    if STATE["journals"]:
        lines.append("Journals: " + ", ".join(sorted(STATE["journals"])))
    lines.append(f"Years: {STATE['years']['min']} - {STATE['years']['max']}")
    lines.append("Paper types: " + ", ".join(sorted(STATE["types"])))
    if STATE["check"]:
        lines.append("Sources-agree filter: " + STATE["check"])
    if STATE["q"]:
        lines.append("Search text: " + STATE["q"])
    return lines


PDF_MAX_PAPERS = 3000                 # keeps the file a sensible size; the first page says when this cut applies
_PDF_KEEP = set("‘’“”–—•…")


def pdf_text(v) -> str:
    """Escape for reportlab and drop characters the bundled font cannot draw (they would print as black boxes)."""
    v = unicodedata.normalize("NFKC", str(v or "")).replace("‐", "-").replace("‑", "-").replace("−", "-")
    v = "".join(ch if (ord(ch) <= 0x17F or ch in _PDF_KEEP) and ch.isprintable() else " " for ch in v)
    return xml_escape(v)


def export_pdf(d: pd.DataFrame, with_abstracts: bool = True) -> Path:
    """Write a PDF of the given papers (most cited first): a first page that lists the filters, then one entry per paper."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer

    fonts = Path(reportlab.__file__).parent / "fonts"            # Bitstream Vera ships inside reportlab
    for name, file in (("Vera", "Vera.ttf"), ("Vera-Bold", "VeraBd.ttf"), ("Vera-Italic", "VeraIt.ttf"), ("Vera-BoldItalic", "VeraBI.ttf")):
        pdfmetrics.registerFont(TTFont(name, str(fonts / file)))
    pdfmetrics.registerFontFamily("Vera", normal="Vera", bold="Vera-Bold", italic="Vera-Italic", boldItalic="Vera-BoldItalic")
    body = ParagraphStyle("body", fontName="Vera", fontSize=8.5, leading=11)
    head = ParagraphStyle("head", parent=body, fontName="Vera-Bold", fontSize=10, leading=12.5, spaceBefore=6)
    small = ParagraphStyle("small", parent=body, textColor=colors.HexColor("#555555"))
    title = ParagraphStyle("title", parent=body, fontName="Vera-Bold", fontSize=18, leading=22, spaceAfter=6)

    total = len(d)
    d = d.sort_values("TC", ascending=False).head(PDF_MAX_PAPERS)
    story = [Paragraph("Journal Explorer - selected papers", title),
             Paragraph(f"{total:,} papers match" + (f" - the {PDF_MAX_PAPERS:,} most cited are listed" if total > PDF_MAX_PAPERS else "") +
                       f" - created {datetime.now():%Y-%m-%d %H:%M}", small), Spacer(1, 6), Paragraph("<b>Filters</b>", body)]
    story += [Paragraph(pdf_text(line), body) for line in describe_filters()]
    story += [Spacer(1, 6), Paragraph("Data: OpenAlex + Crossref (see 'About this data' in the app). Sorted by citations "
                                      "(OpenAlex count).", small), Spacer(1, 10)]
    for i, r in enumerate(d.itertuples(), start=1):
        year = "" if pd.isna(r.PY) else int(r.PY)
        cited = "" if pd.isna(r.TC) else f" - cited {int(r.TC)}"
        bits = [Paragraph(f"{i}. {pdf_text(r.TI)}", head),
                Paragraph(pdf_text(r.AF or r.AU or "") + f" - <i>{pdf_text(r.Journal)}</i> {year}{cited}", small)]
        if r.DI:
            bits.append(Paragraph(f'<link href="https://doi.org/{pdf_text(r.DI)}" color="blue">https://doi.org/{pdf_text(r.DI)}</link>', small))
        if r.DE:
            bits.append(Paragraph("<b>Keywords:</b> " + pdf_text(r.DE), body))
        if with_abstracts and r.AB:
            bits.append(Paragraph(pdf_text(str(r.AB)[:1500]), body))
        story.append(KeepTogether(bits + [Spacer(1, 4)]))

    out = APP_DIR / f"Explore_papers_{datetime.now():%Y%m%d_%H%M}.pdf"
    tmp = out.with_name(f"{out.stem}.building-{uuid.uuid4().hex[:8]}.pdf")
    try:                                                         # temp file first, so a failure never leaves a broken PDF
        SimpleDocTemplate(str(tmp), pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
                          title="Journal Explorer - selected papers").build(story)
        tmp.replace(out)
    finally:
        tmp.unlink(missing_ok=True)
    return out


# ---- downloading: two sources at once, with progress the full-screen overlay can show ----
def progress(abbr: str, source: str, done: int, total: int) -> None:
    STATE["jobs"][abbr][source] = [done, total]


def download_all(journals: list[dict]) -> None:
    """Runs in a background thread so the page stays responsive."""
    STATE.update(busy=True, error="", started=time.time(),
                 jobs={j["abbr"]: {"name": j["name"], "oa": [0, 0], "cr": [0, 0], "state": "waiting"} for j in journals})
    try:
        for j in journals:
            STATE["jobs"][j["abbr"]]["state"] = "downloading"
            refresh_journal(j, progress)
            STATE["jobs"][j["abbr"]]["state"] = "done"
    except Exception as e:                                       # show the problem instead of dying silently
        STATE["error"] = str(e)
    finally:
        STATE["busy"] = False


# ---- the "About this data" text lives above (DISCLAIMER) ----
HERO_STYLE = "background: linear-gradient(135deg, #16385c 0%, #2b6cb0 100%); color: white; border-radius: 16px"


@ui.page("/")
def main_page():
    ui.colors(primary="#1f4e79")
    ui.dark_mode(False)                                           # always light, whatever the OS theme is
    ui.query("body").style("background-color: #f4f6f8")           # explicit page colour (transparent looked black)
    ACTIVE = {"tab": "Papers"}
    # Pop-up editors live under this stable element. A dialog created inside something that refreshes (like the picking
    # block, which redraws on every edit) would be destroyed mid-typing.
    modal_host = ui.element("div")

    # =============================== small helpers ===============================
    def show_about():
        with ui.dialog() as dlg, ui.card().classes("w-[800px] max-w-full"):
            ui.markdown(DISCLAIMER).classes("max-h-[70vh] overflow-auto")
            with ui.row():
                ui.button("I understand", on_click=dlg.close)
                if DF.empty:                                      # first run: offer the download straight away
                    ui.button("I understand - download the journals now", icon="cloud_download",
                              on_click=lambda: (dlg.close(), run_download(list(CONFIG["journals"])))).props("color=primary")
        dlg.open()

    def refresh_views():
        """Redraw whatever is on screen after a filter changed (only the visible tab, to stay quick)."""
        chain_chips.refresh()
        chain_buttons.refresh()
        if ACTIVE["tab"] == "Papers":
            papers.refresh()
        elif ACTIVE["tab"] == "Explore":
            if EX["dirty"]:
                ex_shuffle()
            hero.refresh(); charts.refresh(); plist.refresh()

    def changed():                                                # any filter changed
        EX["dirty"] = True
        refresh_views()

    def reload_data():
        global DF, YB
        was_full = (STATE["years"]["min"], STATE["years"]["max"]) == YB
        DF = load_all()
        YB = year_bounds()
        if was_full:
            STATE["years"] = {"min": YB[0], "max": YB[1]}
        EX["dirty"] = True
        empty_state.refresh(); filters.refresh(); papers_filters.refresh(); journals_panel.refresh(); refresh_views()

    # =============================== full-screen download overlay ===============================
    with ui.dialog().props("persistent maximized") as dl_dialog:
        with ui.card().classes("w-full h-full items-center justify-center gap-6").style("background:#f4f6f8"):
            ui.label("Downloading your journals").classes("text-h3 text-primary")
            ui.label("OpenAlex (the papers) and Crossref (the reference lists) are downloaded at the same time. "
                     "Keep this window open - it closes by itself when done.").classes("text-subtitle1 text-grey-8")
            overall_bar = ui.linear_progress(value=0, show_value=False, size="30px").classes("w-3/4").props("rounded instant-feedback")
            overall_lbl = ui.label().classes("text-h5")
            elapsed_lbl = ui.label().classes("text-caption text-grey-7")

            @ui.refreshable
            def job_rows():
                with ui.column().classes("w-3/4 gap-3"):
                    for abbr, job in STATE["jobs"].items():
                        with ui.card().classes("w-full").props("flat bordered"):
                            with ui.row().classes("items-center w-full"):
                                ui.label(job["name"]).classes("text-subtitle1 text-bold")
                                ui.space()
                                ui.label({"waiting": "waiting", "downloading": "downloading", "done": "done ✓"}[job["state"]]).classes("text-caption")
                            for key, label in (("oa", "OpenAlex - papers"), ("cr", "Crossref - references")):
                                done, total = job[key]
                                with ui.row().classes("items-center w-full no-wrap"):
                                    ui.label(label).classes("w-48 text-caption")
                                    ui.linear_progress(value=(done / total) if total else 0, show_value=False).classes("grow")
                                    ui.label(f"{done:,} / {total:,}" if total else "starting...").classes("w-36 text-caption text-right")
            job_rows()

    def tick():                                                   # twice a second while a download runs
        if not STATE["busy"]:
            return
        fr = [((j["oa"][0] / j["oa"][1] if j["oa"][1] else 0) + (j["cr"][0] / j["cr"][1] if j["cr"][1] else 0)) / 2
              for j in STATE["jobs"].values()]
        pct = sum(fr) / max(len(fr), 1)
        overall_bar.value = pct
        overall_lbl.text = f"{pct * 100:.0f}%"
        s = int(time.time() - STATE["started"])
        elapsed_lbl.text = f"elapsed {s // 60}:{s % 60:02d}"
        job_rows.refresh()
    ui.timer(0.5, tick)

    async def run_download(journals: list[dict]):
        if STATE["busy"]:
            return
        dl_dialog.open()
        await run.io_bound(download_all, journals)
        dl_dialog.close()
        if STATE["error"]:
            ui.notify(f"Download problem: {STATE['error']}", type="negative", multi_line=True, close_button=True, timeout=0)
        else:
            ui.notify("Download finished", type="positive")
        reload_data()

    async def do_pdf():
        d = filtered()
        if d.empty:
            ui.notify("No papers to export - the filters leave nothing.", type="warning")
            return
        note = ui.notification("Building the PDF...", spinner=True, timeout=None)
        try:
            path = await run.io_bound(export_pdf, d, EX["pdf_abs"])
        except Exception as e:
            note.dismiss()
            ui.notify(f"PDF export failed: {e}", type="negative", multi_line=True, close_button=True, timeout=0)
            return
        note.dismiss()
        ui.download(path)
        ui.notify(f"Saved {path}", type="positive")

    async def do_selection_excel():
        d = filtered()
        if d.empty:
            ui.notify("No papers to export - the filters leave nothing.", type="warning")
            return
        note = ui.notification("Building the Excel file...", spinner=True, timeout=None)
        try:
            path = await run.io_bound(export_excel, d)
        except Exception as e:
            note.dismiss()
            ui.notify(f"Excel export failed: {e}", type="negative", multi_line=True, close_button=True, timeout=0)
            return
        note.dismiss()
        ui.download(path)
        ui.notify(f"Saved {path}", type="positive")

    # =============================== chain filters (AND / OR / AND NOT, any length) ===============================
    def cycle_op(c: dict):
        c["op"] = OPS[(OPS.index(c["op"]) + 1) % len(OPS)]
        changed()

    def remove_term(field: str, i: int):
        STATE["chains"][field].pop(i)
        changed()

    def add_term(field: str, text: str, op: str | None = None):
        text = (text or "").strip()
        if text:
            STATE["chains"][field].append({"op": (op or "AND") if STATE["chains"][field] else "AND", "text": text})
            changed()

    def render_chain(field: str, big: bool = False):
        """The terms of one chain as buttons: click the AND/OR word to switch it, click a term to remove it."""
        size = "size=lg" if big else "size=sm"
        for i, c in enumerate(STATE["chains"][field]):
            if i > 0:
                ui.button(c["op"], on_click=lambda c=c: cycle_op(c)).props(f"flat dense no-caps {size}" + (" color=white" if big else ""))
            label = ("NOT " if i == 0 and c["op"] == "AND NOT" else "") + c["text"]
            ui.button(f"{label}  ✕", on_click=lambda f=field, i=i: remove_term(f, i)).props(
                "unelevated no-caps rounded color=amber-8 text-color=black " + size if big else f"outline no-caps rounded dense {size}")

    @ui.refreshable
    def chain_chips(field: str):
        render_chain(field)

    def open_chain_modal(field: str):
        """A list editor for one chain: edit any term, change its AND/OR, delete it, add more - no length limit."""
        chain = STATE["chains"][field]
        modal_host.clear()
        with modal_host, ui.dialog() as dlg, ui.card().classes("w-[680px] max-w-full"):
            ui.label(f"{CHAIN_LABELS[field]} filter").classes("text-h6")
            ui.label("Terms are joined left to right; AND binds tighter than OR. Changes apply immediately.").classes("text-caption")

            @ui.refreshable
            def rows():
                with ui.column().classes("w-full gap-1"):
                    for i, c in enumerate(chain):
                        with ui.row().classes("items-center w-full no-wrap"):
                            opts = {"AND": "start", "AND NOT": "NOT"} if i == 0 else {o: o for o in OPS}
                            ui.select(opts, value=c["op"], on_change=lambda e, c=c: (c.update(op=e.value), changed())).props("dense outlined").classes("w-28")
                            ui.input(value=c["text"], on_change=lambda e, c=c: (c.update(text=e.value or ""), changed())).props(
                                "dense outlined debounce=500").classes("grow")
                            ui.button(icon="delete", on_click=lambda i=i: (chain.pop(i), rows.refresh(), changed())).props("flat round color=negative")
                    if not chain:
                        ui.label("No terms yet.").classes("text-caption")
            rows()
            with ui.row():
                ui.button("Add term", icon="add", on_click=lambda: (chain.append({"op": "AND", "text": ""}), rows.refresh())).props("flat")
                ui.button("Clear all", on_click=lambda: (chain.clear(), rows.refresh(), changed())).props("flat color=negative")
                ui.button("Done", on_click=dlg.close).props("color=primary")
        dlg.open()

    def chain_widget(field: str):
        with ui.row().classes("items-center gap-1 no-wrap"):
            ui.label(CHAIN_LABELS[field] + ":").classes("text-caption text-bold")
            chain_chips(field)
            box = ui.input(placeholder="add term + Enter").props("dense outlined").classes("w-40")

            def add():
                add_term(field, box.value, EX["next_op"])
                box.value = ""
            box.on("keydown.enter", add)
            ui.button(icon="edit", on_click=lambda f=field: open_chain_modal(f)).props("flat round dense").tooltip("Edit the whole list")

    # =============================== the shared filter block (own visual card) ===============================
    @ui.refreshable
    def filters(show_keywords: bool):
        with ui.card().classes("w-full gap-2").props("flat bordered"):
            ui.label("Filters").classes("text-subtitle1 text-bold")
            with ui.row().classes("items-center gap-4 w-full"):
                ui.input("Search all text", value=STATE["q"], on_change=lambda e: (STATE.update(q=e.value or ""), changed())).props(
                    "clearable outlined dense debounce=400").classes("w-64")
                ui.select({j["abbr"]: j["name"] for j in CONFIG["journals"]}, multiple=True, value=sorted(STATE["journals"]),
                          label="Journals (all)", on_change=lambda e: (STATE.update(journals=set(e.value)), changed())
                          ).props("outlined dense use-chips").classes("w-72")
                with ui.column().classes("gap-0 w-64"):
                    ui.label("Years").classes("text-caption")
                    ui.range(min=YB[0], max=YB[1], value=dict(STATE["years"]),
                             on_change=lambda e: (STATE.update(years=dict(e.value)), changed())).props("label-always")
                ui.select({"": "All records", "✅": "✅ Both sources agree", "⚠": "⚠️ Sources conflict",
                           "◻": "◻ Only one source"}, value=STATE["check"], label="Sources agree?",
                          on_change=lambda e: (STATE.update(check=e.value), changed())).props("outlined dense").classes("w-52")
            with ui.row().classes("items-center gap-1"):
                ui.label("Paper types:").classes("text-caption")
                for t in DOC_TYPES:
                    def toggle(e, t=t):
                        STATE["types"].add(t) if e.value else STATE["types"].discard(t)
                        changed()
                    ui.checkbox(t, value=t in STATE["types"], on_change=toggle)
            with ui.row().classes("items-center gap-6"):
                for field in CHAIN_LABELS:
                    if show_keywords or field != "keywords":
                        chain_widget(field)

    # =============================== Papers tab filters: as they always were ===============================
    @ui.refreshable
    def papers_filters():
        """Search, journals, confidence and paper types - the original Papers tab controls (shared state, so what you set
        in Explore shows here too)."""
        with ui.row().classes("items-center gap-4 w-full"):
            ui.input("Search title, abstract, keywords", value=STATE["q"], on_change=lambda e: (STATE.update(q=e.value or ""), changed())
                     ).props("clearable outlined dense debounce=400").classes("w-96")
            ui.select({j["abbr"]: j["name"] for j in CONFIG["journals"]}, multiple=True, value=sorted(STATE["journals"]),
                      label="Journals (all)", on_change=lambda e: (STATE.update(journals=set(e.value)), changed())
                      ).props("outlined dense use-chips").classes("w-72")
        ui.select({"": "All records", "\u2705": "\u2705 Both sources agree", "\u26a0": "\u26a0\ufe0f Sources conflict",
                   "\u25fb": "\u25fb Only one source"}, value=STATE["check"], label="Confidence",
                  on_change=lambda e: (STATE.update(check=e.value), changed())).props("outlined dense").classes("w-56")
        with ui.row().classes("items-center gap-1"):
            ui.label("Show:").classes("text-caption")
            for t in DOC_TYPES:
                def toggle(e, t=t):
                    STATE["types"].add(t) if e.value else STATE["types"].discard(t)
                    changed()
                ui.checkbox(t, value=t in STATE["types"], on_change=toggle)

    def reset_years():
        STATE["years"] = {"min": YB[0], "max": YB[1]}
        filters.refresh()
        changed()

    @ui.refreshable
    def chain_buttons():
        """One button per text field. Conditions (AND / OR / AND NOT, any number) are written and edited in the pop-up,
        so nothing grows on the page. The number shows how many conditions are active."""
        with ui.row().classes("items-center gap-2"):
            ui.label("Conditions:").classes("text-caption")
            for field, label in CHAIN_LABELS.items():
                n = len([c for c in STATE["chains"][field] if c["text"].strip()])
                ui.button(label + (f"  ({n})" if n else ""), icon="tune", on_click=lambda f=field: open_chain_modal(f)).props(
                    "outline no-caps dense " + ("color=primary" if n else "color=grey-7"))
            if (STATE["years"]["min"], STATE["years"]["max"]) != YB:   # a years range set in Explore also applies here
                ui.button(f'Years {STATE["years"]["min"]}-{STATE["years"]["max"]}  \u2715', on_click=reset_years).props(
                    "outline no-caps dense color=amber-9").tooltip("Set in the Explore tab - click to clear")

    # =============================== the paper detail dialog ===============================
    def show_paper(r: dict):
        with ui.dialog() as dlg, ui.card().classes("w-[900px] max-w-full"):
            ui.label(r["TI"]).classes("text-h6")
            ui.label(f'{r.get("AF") or ""} - {r["Journal"]} {r["PY"]}').classes("text-caption")
            if r.get("Conflicts"):
                ui.markdown(f'**Where OpenAlex and Crossref disagree:** {r["Conflicts"]}')
            ui.markdown(f'**Abstract**\n\n{r.get("AB") or "_none_"}')
            ui.markdown(f'**Cited references**\n\n```\n{r.get("CR") or "none"}\n```').classes("max-h-64 overflow-auto")
            ui.link("Open at publisher", f'https://doi.org/{r["DI"]}', new_tab=True)
            ui.button("Close", on_click=dlg.close)
        dlg.open()

    # =============================== Papers tab: the big table ===============================
    @ui.refreshable
    def papers():
        d = filtered()
        ui.label(f"{len(d):,} papers").classes("text-subtitle2")
        cols = [("Journal", 90), ("PY", 80), ("TI", 520), ("AU", 240), ("Type", 130), ("Check", 210), ("TC", 90), ("CitedCrossref", 110), ("DE", 300), ("DI", 180)]
        names = {"PY": "Year", "TI": "Title", "AU": "Authors", "TC": "Cited (OpenAlex)", "CitedCrossref": "Cited (Crossref)",
                 "DE": "Keywords", "DI": "DOI", "Check": "Sources agree?"}
        grid = ui.aggrid({
            "columnDefs": [{"field": c, "headerName": names.get(c, c), "minWidth": w, "flex": w} for c, w in cols],
            "rowData": d[[c for c, _ in cols] + ["AB", "CR", "AF", "Conflicts"]].head(50000).where(d.notna(), None).to_dict("records"),
            "defaultColDef": {"sortable": True, "filter": True, "resizable": True},
            "pagination": True, "paginationPageSize": 50,
        }).classes("w-full").style("height: 70vh")
        grid.on("cellClicked", lambda e: show_paper(e.args["data"]))

    # =============================== Explore tab: the keyword slot machine ===============================
    def ex_counts() -> collections.Counter:
        """Keywords of the papers that match everything right now, minus the ones already in the chain."""
        d = filtered()
        inchain = {c["text"].strip().lower() for c in STATE["chains"]["keywords"]}
        return collections.Counter(k for kws in d["_kw"] for k in kws if k.lower() not in inchain) if not d.empty else collections.Counter()

    def ex_shuffle():
        """Draw N_PICKS keywords. 'Surprise' 0 = weighted to the common ones, 1 = every keyword equally likely."""
        EX["dirty"] = False
        items = [(k, n) for k, n in ex_counts().items() if n >= EX["min_papers"]]
        if not items:
            EX["picks"] = []
            return
        w = np.array([n ** (1 - EX["surprise"]) for _, n in items], dtype=float)
        idx = np.random.choice(len(items), size=min(N_PICKS, len(items)), replace=False, p=w / w.sum())
        EX["picks"] = [items[i] for i in idx]

    def ex_pick(word: str):
        add_term("keywords", word, EX["next_op"])

    def ex_random_start():
        pool = [k for k, n in ex_counts().items() if n >= max(EX["min_papers"], 5)]
        if pool:
            add_term("keywords", random.choice(pool))

    @ui.refreshable
    def hero():
        """The centre of the Explore tab: the words you have picked and the words on offer."""
        with ui.card().classes("w-full q-pa-lg gap-4").style(HERO_STYLE):
            if DF.empty:
                ui.label("No data yet").classes("text-h5")
                ui.button("Download the journals", icon="cloud_download", on_click=lambda: run_download(list(CONFIG["journals"]))).props("color=white text-color=primary")
                return
            d = filtered()
            with ui.row().classes("items-center w-full gap-4"):
                ui.label("Pick your words").classes("text-h4")
                ui.space()
                ui.label(f"{len(d):,} papers match").classes("text-h6")
                ui.button("Export these papers to Excel", icon="table_view", on_click=do_selection_excel).props("unelevated no-caps color=white text-color=primary")
                ui.button("PDF", icon="picture_as_pdf", on_click=do_pdf).props("unelevated no-caps color=white text-color=primary")
                ui.checkbox("with abstracts", value=EX["pdf_abs"], on_change=lambda e: EX.update(pdf_abs=e.value)).props("dark")
            # --- what you have picked ---
            with ui.row().classes("items-center gap-2"):
                if STATE["chains"]["keywords"]:
                    render_chain("keywords", big=True)
                    ui.button("Edit list", icon="edit", on_click=lambda: open_chain_modal("keywords")).props("flat no-caps color=white")
                    ui.button("Clear", on_click=lambda: (STATE["chains"]["keywords"].clear(), changed())).props("flat no-caps color=white")
                else:
                    ui.label("Nothing picked yet - choose a word below, type your own, or").classes("text-subtitle1")
                    ui.button("Random start", icon="casino", on_click=ex_random_start).props("unelevated no-caps color=amber-8 text-color=black")
            ui.separator().props("dark")
            # --- what is on offer ---
            with ui.row().classes("items-center gap-3"):
                ui.button("Shuffle", icon="shuffle", on_click=lambda: (ex_shuffle(), hero.refresh())).props("unelevated color=amber-8 text-color=black size=lg")
                ui.label("Next word joins with:").classes("text-subtitle1")
                ui.toggle({"AND": "AND", "OR": "OR", "AND NOT": "NOT"}, value=EX["next_op"],
                          on_change=lambda e: EX.update(next_op=e.value)).props("toggle-color=amber-8 toggle-text-color=black color=white text-color=primary")
            if not EX["picks"]:
                ex_shuffle()
            with ui.row().classes("gap-3 w-full"):
                for k, n in EX["picks"]:
                    with ui.button(on_click=lambda k=k: ex_pick(k)).props("unelevated color=white text-color=primary no-caps size=lg"):
                        ui.label(k).classes("text-h6")
                        ui.badge(f"{n}", color="primary").classes("q-ml-sm")
                if not EX["picks"]:
                    ui.label("No more keywords with these filters - remove a word or lower the 'at least N papers' number.").classes("text-subtitle1")
            with ui.row().classes("items-center gap-6"):
                box = ui.input(placeholder="or type your own keyword + Enter").props("dense outlined dark").classes("w-72")
                box.on("keydown.enter", lambda: (add_term("keywords", box.value, EX["next_op"]), setattr(box, "value", "")))
                with ui.column().classes("gap-0 w-56"):
                    ui.label("Words offered: common  <->  surprising").classes("text-caption")
                    ui.slider(min=0, max=1, step=0.1, value=EX["surprise"], on_change=lambda e: EX.update(surprise=e.value)).props("color=amber-8 dark")
                ui.number("Only words in at least N papers", value=EX["min_papers"], min=1, max=50, step=1,
                          on_change=lambda e: EX.update(min_papers=int(e.value or 1))).props("dense outlined dark").classes("w-64")

    @ui.refreshable
    def charts():
        """Papers over time for every picked keyword (or per journal when nothing is picked), and the most common other keywords."""
        d0 = filtered(skip=("keywords",))                          # everything except the keyword chain
        if d0.empty:
            return
        chain = [c for c in STATE["chains"]["keywords"] if c["text"].strip()]
        year = pd.to_numeric(d0["PY"], errors="coerce")
        if year.dropna().empty:
            return
        years = list(range(int(year.min()), int(year.max()) + 1))
        totals = year.value_counts().reindex(years, fill_value=0)
        pct = EX["chart_mode"] == "pct"

        def line(name, mask, **extra):
            counts = year[mask].value_counts().reindex(years, fill_value=0)
            vals = (counts / totals.replace(0, np.nan) * 100).round(1).fillna(0) if pct else counts
            return {"name": name, "type": "line", "smooth": True, "showSymbol": False, "data": [float(v) if pct else int(v) for v in vals], **extra}

        if chain:
            series = [line(c["text"], term_mask(d0, "keywords", c["text"]).to_numpy()) for c in chain if c["op"] != "AND NOT"]
            m = chain_mask(d0, "keywords", chain)
            if len(chain) > 1 and m is not None:
                series.append(line("Papers matching the whole chain", m, lineStyle={"width": 4, "type": "dashed"}))
            title = "Papers per year for your picked keywords"
        else:
            series = [line(j, (d0["Journal"] == j).to_numpy()) for j in sorted(d0["Journal"].unique())]
            title = "Papers per year (pick keywords above to see them here)"
        d = filtered()
        counts = ex_counts()
        top = counts.most_common(20) if EX["kw_mode"] == "most" else sorted(counts.items(), key=lambda kv: kv[1])[:20]
        with ui.row().classes("w-full no-wrap gap-4"):
            ui.echart({"title": {"text": title + (" (% of that year's papers)" if pct else "")}, "tooltip": {"trigger": "axis"},
                       "legend": {"top": 30, "type": "scroll"}, "grid": {"top": 80}, "xAxis": {"type": "category", "data": years},
                       "yAxis": {"type": "value"}, "series": series}).classes("w-1/2").style("height:380px")
            ui.echart({"title": {"text": "Most common other keywords in these papers" if EX["kw_mode"] == "most" else "Rarest other keywords"},
                       "grid": {"left": 190, "top": 50}, "tooltip": {}, "xAxis": {"type": "value"},
                       "yAxis": {"type": "category", "data": [k for k, _ in top][::-1]},
                       "series": [{"type": "bar", "data": [int(n) for _, n in top][::-1]}]}).classes("w-1/2").style("height:380px")

    @ui.refreshable
    def plist():
        d = filtered()
        if d.empty:
            return
        rows = d.sort_values("TC", ascending=False).head(100)
        ui.label(f"The {len(rows)} most cited of the {len(d):,} matching papers").classes("text-subtitle2")
        tbl = ui.table(columns=[{"name": "t", "label": "Title", "field": "t", "align": "left"},
                                {"name": "j", "label": "Journal", "field": "j", "sortable": True},
                                {"name": "y", "label": "Year", "field": "y", "sortable": True},
                                {"name": "c", "label": "Cited", "field": "c", "sortable": True}],
                       rows=[{"t": r.TI, "j": r.Journal, "y": None if pd.isna(r.PY) else int(r.PY), "c": None if pd.isna(r.TC) else int(r.TC),
                              "i": int(i)} for i, r in zip(range(len(rows)), rows.itertuples())],
                       row_key="i", pagination=10).classes("w-full")
        by_index = rows.reset_index(drop=True)
        tbl.on("rowClick", lambda e: show_paper(by_index.iloc[e.args[1]["i"]].fillna("").to_dict()))

    # =============================== header ===============================
    with ui.header().classes("items-center gap-4"):
        ui.label("Journal Explorer").classes("text-h6")
        ui.space()
        ui.button("Update data", icon="sync", on_click=lambda: run_download(list(CONFIG["journals"]))).props("flat color=white")

        async def do_export():
            if EXPORT_LOCK.locked():                              # a second click must not disturb the running export
                ui.notify("An export is already running - please wait for it to finish.", type="warning")
                return
            export_btn.props("loading disable")                   # spinner + no second click while it runs
            note = ui.notification("Building the Excel file - this takes a moment. Please wait...", spinner=True, timeout=None)
            try:
                path = await run.io_bound(export_excel)
            except Exception as e:                                # tell the user; the old Journals.xlsx is untouched
                note.dismiss()
                ui.notify(f"Excel export failed: {e}", type="negative", multi_line=True, close_button=True, timeout=0)
                return
            finally:
                export_btn.props(remove="loading disable")
            note.dismiss()
            ui.download(path)
            ui.notify(f"Saved {path}", type="positive")

        export_btn = ui.button("Export to Excel", icon="download", on_click=do_export).props("flat color=white")
        ui.button("About this data", icon="info", on_click=show_about).props("flat color=white")

    @ui.refreshable
    def empty_state():
        """Nothing downloaded yet: one big, centred button - impossible to miss on first start."""
        if not DF.empty:
            return
        with ui.card().classes("w-full items-center q-pa-xl gap-4").props("flat bordered"):
            ui.icon("cloud_download", size="96px", color="primary")
            ui.label("No papers downloaded yet").classes("text-h4")
            ui.label("One click fetches all your journals (a few minutes, needs internet).").classes("text-subtitle1 text-grey-8")
            ui.button("Download papers", icon="cloud_download", on_click=lambda: run_download(list(CONFIG["journals"]))).props(
                "size=xl unelevated rounded color=primary").classes("text-h5 q-px-xl q-py-md")
            ui.button("Please read 'About this data' first", icon="info", on_click=show_about).props("flat no-caps")
    empty_state()

    # =============================== tabs ===============================
    def on_tab(e):
        ACTIVE["tab"] = e.value
        EX["dirty"] = True
        filters.refresh(); papers_filters.refresh()               # both tabs re-read the shared state
        refresh_views()

    with ui.tabs().classes("w-full") as tabs:
        ui.tab("Papers", icon="table_rows")
        ui.tab("Explore", icon="casino")
        ui.tab("Journals", icon="settings")

    with ui.tab_panels(tabs, value="Papers", on_change=on_tab).classes("w-full").style("background: transparent"):
        with ui.tab_panel("Papers"):
            papers_filters()
            chain_buttons()
            papers()

        with ui.tab_panel("Explore"):
            hero()
            filters(False)
            with ui.row().classes("items-center gap-4"):
                ui.toggle({"count": "Papers", "pct": "% of that year"}, value=EX["chart_mode"],
                          on_change=lambda e: (EX.update(chart_mode=e.value), charts.refresh()))
                ui.toggle({"most": "Most used keywords", "least": "Rarest keywords"}, value=EX["kw_mode"],
                          on_change=lambda e: (EX.update(kw_mode=e.value), charts.refresh()))
            charts()
            plist()

        with ui.tab_panel("Journals"):
            ui.label("Tracked journals").classes("text-h6")

            @ui.refreshable
            def journals_panel():
                for j in CONFIG["journals"]:
                    n_dl = len(pd.read_pickle(pkl(j))) if pkl(j).exists() else 0
                    with ui.row().classes("items-center w-full border-b"):
                        ui.label(j["name"]).classes("w-96 text-bold")
                        ui.label(f'sheet "{j["abbr"]}"').classes("w-32 text-caption")
                        ui.label(f"{n_dl:,} downloaded" + (f' of ~{j["works_count"]:,} listed' if j.get("works_count") else "")).classes("w-64")
                        ui.label("last update: " + CONFIG["last_fetch"].get(j["abbr"], "never")).classes("w-48 text-caption")
                        ui.button(icon="delete", on_click=lambda j=j: remove(j)).props("flat color=negative")

            def remove(j):
                CONFIG["journals"].remove(j)
                pkl(j).unlink(missing_ok=True)
                STATE["journals"].discard(j["abbr"])
                save_config(CONFIG)
                reload_data()

            journals_panel()

            ui.separator()
            ui.label("Add a journal").classes("text-h6")
            box = ui.input("Type a journal name...", on_change=lambda e: search(e.value)).props("outlined dense debounce=350").classes("w-96")
            results = ui.column()

            def search(text: str):                              # OpenAlex type-ahead over ~250,000 journals
                results.clear()
                if len(text) < 3:
                    return
                with results:
                    for s in Sources().autocomplete(text)[:8]:
                        ui.button(f'{s["display_name"]}  -  {s.get("hint") or ""}  ({s["works_count"]:,} papers)',
                                  on_click=lambda s=s: add(s)).props("flat no-caps align=left").classes("w-full")

            async def add(s):
                abbr = "".join(w[0] for w in re.findall(r"[A-Za-z]+", s["display_name"]) if w[0].isupper())[:8]
                abbr = abbr or re.sub(r"[^A-Za-z0-9]", "", s["display_name"])[:8] or "J"
                while any(abbr.lower() == j["abbr"].lower() for j in CONFIG["journals"]) or abbr.lower() in ("summary", "history"):
                    abbr += "2"
                j = {"name": s["display_name"], "abbr": abbr, "oa_id": s["id"].split("/")[-1],
                     "issn": s["external_id"], "works_count": s["works_count"]}
                CONFIG["journals"].append(j)
                save_config(CONFIG)
                box.value, _ = "", results.clear()
                journals_panel.refresh()
                await run_download([j])


# port fixed so a bookmark keeps working; reload=False is required when running as an exe
if __name__ in {"__main__", "__mp_main__"}:
    ui.run(title="Journal Explorer", port=8765, reload=False, show=not __import__("os").environ.get("NO_BROWSER"))
