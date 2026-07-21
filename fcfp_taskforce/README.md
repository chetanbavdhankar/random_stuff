# Global FCFP AI Productivity — Use Case Hub

Two ways to run. **Backend mode is the one you want** — it edits the real
Excel file in place and keeps all formatting, dropdowns and colours.

---

## Backend mode (recommended) — edits the same file, formatting preserved

A tiny local Python server opens the real `.xlsx` with **openpyxl** and
writes only the new row, so conditional formatting (RAG colours),
data-validation dropdowns, column widths, merged cells and cell styles all
stay intact. New rows inherit the style of the row above them, and dropdown
ranges are extended to cover them.

### One-time setup
1. Install **Python 3.9+** (Windows: python.org or Microsoft Store).
2. Put these three files in the same folder:
   `server.py`, `requirements.txt`, `fcfp-ai-usecase-hub.html`.
3. Install dependencies — in PowerShell / Terminal in that folder:
   ```
   pip install -r requirements.txt
   ```

### Point it at your tracker
Open `server.py` and set `TRACKER_PATH` near the top, e.g.
```python
TRACKER_PATH = r"C:\Users\you\ING\Global FCFP\UseCaseTracker.xlsx"
SHEET_NAME   = ""     # "" = auto-detect the use-case sheet
```
If that path is inside a **synced SharePoint / OneDrive folder**, every save
syncs back up to SharePoint automatically — that is the practical way to get
"everyone sees it" without any Graph/API setup.

You can also set it without editing the file:
```
# Windows PowerShell
$env:TRACKER_PATH="C:\path\UseCaseTracker.xlsx"; python server.py
# macOS / Linux
TRACKER_PATH="/path/UseCaseTracker.xlsx" python server.py
```

### Run it
```
python server.py
```
Open the address it prints (default `http://127.0.0.1:5000`) in **Edge or
Chrome**. The page auto-detects the backend — no manual file loading. The
banner shows **● Live-editing <file>** when connected. Submitting a use case
writes straight into the file; if the file is open in Excel it tells you to
close it first (so nothing is corrupted).

**Team use:** run this on one shared machine/VM and share the URL, or each
person runs it locally against the same synced SharePoint copy. For a fully
hosted, multi-user version the clean path is a SharePoint list + Power
Automate, or Microsoft Graph with an Entra app registration — `read_tracker`
and `append_usecase` in `server.py` are the two functions to repoint.

---

## Standalone mode (no Python) — fallback

Just open `fcfp-ai-usecase-hub.html` in a browser. It works, but because a
browser can only *rebuild* an `.xlsx`, saving this way **loses** conditional
formatting, dropdowns and other Excel UI. Use this only for a quick look or
when Python isn't available. In Chrome/Edge it can still write back to the
same file you opened (File System Access API); elsewhere it downloads an
updated copy.

---

## Configuring columns / sheet name

Column detection is automatic and fuzzy: your tracker can use different
wording ("UC No", "Use Case Title", "Timeline RAG"…) and a different column
order. The hub reports how many of the 38 fields it matched and lists any it
couldn't find. Unmatched columns from your file are preserved and shown
under "Additional columns from your file".

To force a specific sheet, set `SHEET_NAME` in `server.py` (backend) or
`CONFIG.SHEET_NAME` in the HTML (standalone).
