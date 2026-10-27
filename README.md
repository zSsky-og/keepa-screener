# Keepa screener

Screens a supplier price list against Keepa and tells you which products are worth selling on Amazon:
Amazon on the listing, monthly sales, sales rank, offer count, ROI, Buy Box suppression and HAZMAT.

## Start it

| Computer    | What to do                         |
| ----------- | ---------------------------------- |
| **Windows** | Double-click **`start.bat`**       |
| **Mac**     | Double-click **`start.command`**.  |

The first start installs what it needs . After that the screener opens in your browser at
`http://127.0.0.1:8765/`. Keep the black window open while you work; close it to quit.

If Python is missing, the launcher opens the Python download page. On Windows, tick **"Add python.exe to PATH"** in the installer.

## Using it

1. **New screening** → choose the supplier file (.csv, .xlsx, .xls). Columns are detected and remembered for files with the same headers.
2. Enter your Keepa API key (kept in this browser only, never saved with a run) and press **Check and estimate**.
3. Press **Start**. Results appear live, newest at the top, one product at a time.
4. **Pause** / **Resume** at any time (or press **P**). While paused no Keepa requests are sent and no tokens are used.
5. **Stop** ends the run and keeps everything done so far. The report lists the unchecked rows as `NOT_CHECKED`.
6. **Download report** gives an Excel file with Summary, Passed, Failed and All sheets.

How it talks to Keepa: one product per request. Each result is scored and shown before the next request goes out.
It only waits when your token balance is zero or below, and shows a countdown until the refill.
Lookups are cached for 24 hours, so running the same list again does not cost tokens.

## Command line

```
python keepa_screener.py list.xlsx --api-key KEY      screen a file
python keepa_screener.py list.xlsx --dry-run          show lookups and token cost only
python keepa_screener.py --export-cache               write every cached lookup to keepa_cache_export.csv
```
Ctrl+C once stops after the current product and still saves the report.

## Files

| File                          | What it is                                             |
| ----------------------------- | ------------------------------------------------------ |
| `keepa_screener.py`           | Screening engine, also a command line tool             |
| `app.py`                      | Local web app.                                         |
| `ui.html`                     | The whole interface                                    |
| `start.bat` / `start.command` | Double-click launchers for Windows / Mac               |
| `data/`                       | Created on first start: uploads and one folder per run |
| `.keepa_cache/`               | Cached Keepa lookups                                   |

To keep an old cache and column mappings, copy the `.keepa_cache` folder and `vendor_profiles.json` from the old folder.
