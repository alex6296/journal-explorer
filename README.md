# Journal Explorer

Download every paper of the academic journals you pick, explore them by keyword, and export to Excel or PDF.
Runs on your own computer in a browser tab. Free data sources, no login, no API keys.

**Repository:** https://github.com/alex6296/journal-explorer

> **Open source, not actively maintained.** This is released under the [MIT license](LICENSE): use it, change it,
> fork it. It was built for one person's research and is shared as-is. Issues and pull requests may not get an answer.

## What it does

- **Download** the papers of any journal (search ~250,000 journals by name) from **OpenAlex** and **Crossref**, both at the
  same time, with a full-screen progress bar. The four journals it starts with are *Strategic Management Journal*,
  *Long Range Planning*, *Strategic Organization* and *Strategy Science*.
- **Cross-check** the two sources paper by paper. Each paper is marked: both agree, they conflict (and on which field), or
  only one source has it.
- **Papers tab:** search, sort, and filter by journal, years, paper type and source agreement. Text filters (title,
  authors, keywords, abstract) take **any number of conditions** - Contains, Does not contain, Equals, Does not equal, Begins with,
  Ends with, Blank, Not blank - joined by AND / OR and edited in a pop-up list.
- **Explore tab:** a keyword "slot machine" for finding ideas. Lock a keyword, shuffle to get 10 keywords from papers that
  also have it, lock another, and so on. Charts show the locked keywords over time and the most common other keywords.
  **Export the papers left after all filters to Excel** (same workbook layout as the full export, plus a Filters sheet) **or PDF.**
- **Export to Excel** (header button = everything you downloaded) in the layout of a Web of Science export: one sheet per journal with the same column tags, and a
  Summary sheet with keyword counts.
- **Settings tab (optional):** paste your own [Elsevier/Scopus API key](https://dev.elsevier.com/apikey/manage), with a short
  built-in video on where to get one, to also cross-check every paper against Scopus and pull in its citation count.
  Pasting a key **tests it automatically** (no button to remember) - a verified key IS the switch: it goes straight
  into re-downloading your journals with Scopus included, no extra click, and the button relabels itself "Update
  data (with Scopus)" from then on. A bad or gated key can't turn it on - it would otherwise just mark every paper
  "not in Scopus". Pulls use Scopus Search's documented `start`/`count` paging (capped at 5,000 results per journal -
  that's the API's own limit, not this app's), paced to Elsevier's own guidance of about one request per second, and
  stop themselves well before your weekly quota would run out; a network hiccup on any source gets a couple of
  automatic retries before giving up. A key only returns full results from your institution's network or its VPN;
  the app explains this in a pop-up if a request comes back gated. The whole Scopus path is logged to the console
  (timestamped, key masked) if you want to see what it's doing. **Elsevier's terms make data pulled with your key
  personal to you** - don't hand the `data` folder or an export with Scopus columns filled in to anyone else.

## Run it

**Windows, no install:** download [`JournalExplorer.exe`](https://github.com/alex6296/journal-explorer/raw/main/JournalExplorer.exe)
(root of this repo, 79 MB) and double-click it. Press **Download papers** on first start. Windows may show "More info -> Run anyway"
because the exe is not code-signed.

**Mac (Apple Silicon - M1/M2/M3/M4, not Intel Macs):** download both
[`JournalExplorer-mac`](https://github.com/alex6296/journal-explorer/raw/main/JournalExplorer-mac) and
[`Start JournalExplorer (Mac).command`](<https://github.com/alex6296/journal-explorer/raw/main/Start JournalExplorer (Mac).command>)
into the same folder, then double-click the `.command` file. The first time, macOS will refuse to open it as an
"unidentified developer" - right-click it (or `JournalExplorer-mac`) and choose **Open** once, confirm, and it runs
normally after that.

Both builds are rebuilt by hand (or via `.github/workflows/build.yml` on GitHub Actions), so they can lag behind the source.

**From source** (Windows, Mac or Linux, Python 3.11+):

```bash
pip install -r requirements.txt
python journal_explorer.py
```

A browser tab opens at http://localhost:8765. Downloaded data is saved in a `data` folder next to the app, so the next start is instant.

**Build a standalone app** (PyInstaller cannot cross-compile, so build on the OS you want):

```bash
pip install -r requirements.txt
nicegui-pack --onefile --name JournalExplorer journal_explorer.py
```

`.github/workflows/build.yml` does this for Windows and Mac on GitHub's runners (Actions tab -> Build -> Run workflow).
The Mac app is not code-signed: right-click -> Open the first time.

## Read this before trusting the numbers

The app uses two free, automated sources instead of a curated database like Scopus or Web of Science. In a 250-paper
check against scite (September 2026), papers where both sources agreed matched almost perfectly; where they conflicted,
Crossref was right (OpenAlex often holds the online-first year instead of the print year). Completeness depends on era and
publisher: for example abstracts are missing for most of *Long Range Planning*, and old papers have fewer reference lists.
Keywords are OpenAlex's machine-generated topics, not author keywords. Citation counts will not match Scopus or Web of
Science. The full list of limits is in the app under **About this data** (and in `DISCLAIMER` in `journal_explorer.py`).

## How it is built

One file, [`journal_explorer.py`](journal_explorer.py), commented section by section. It leans on libraries instead of
custom code: [NiceGUI](https://nicegui.io) (the whole UI, charts and tables), [pyalex](https://github.com/J535D165/pyalex)
(OpenAlex), pandas, xlsxwriter (Excel, with calculated values stored beside the formulas) and reportlab (PDF).

## License

[MIT](LICENSE) - Copyright (c) 2026 Alex. Data comes from [OpenAlex](https://openalex.org) (CC0) and
[Crossref](https://www.crossref.org) (metadata is freely reusable); check their terms if you build on this.
