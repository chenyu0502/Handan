# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

This repo contains one project, **HanDan/** (菡萏咖啡．台股損益存摺 — a personal Taiwan-stock portfolio tracker). It has no package manager or test suite: it's a single self-contained HTML file plus two Python helper scripts. `serve.py` can optionally be packaged into a standalone Windows `.exe` via PyInstaller (see Commands) so it runs without a Python install, but this is a manual, occasional step — there's still no CI/lint/test pipeline.

- `HanDan/菡萏咖啡-台股損益存摺.html` — the entire application (~2600 lines). All markup, CSS, and JS live in this one file. Uses Chart.js and SheetJS (xlsx) from CDN `<script>` tags; no bundler, no npm.
- `HanDan/fetch_close.py` — CLI tool that scrapes TWSE/TPEx closing prices and can patch the HTML's embedded price block in place.
- `HanDan/serve.py` — a local `http.server`-based server that serves the HTML and exposes `/api/xlsx-prices` (auto-called on page load, reads straight from `菡萏咖啡.xlsx`) and `/api/close-prices` (only called by the "取得盤後收盤價" button, does the live TWSE/TPEx fetch and writes the result back into the xlsx).
- `HanDan/開啟存摺.bat` — double-click launcher; runs `serve.py` from the venv and lets it auto-open the browser. This is the intended way to start the app day-to-day, instead of opening the HTML file directly or typing the `python serve.py` command.
- `HanDan/README.md` — end-user-facing usage doc (Traditional Chinese) for people other than the primary maintainer: how to launch, SmartScreen/firewall prompts, the two accounts, why manual edits live in browser storage (not the xlsx) and need periodic backup export, how price/balance sync works, and 04 記一筆's weekly buy/sell rollup + 匯入 button caveat. Keep it in sync when user-facing behavior changes — it's a separate audience from this file.
- `HanDan/菡萏咖啡.xlsx` — the user's original hand-kept ledger (Excel). Its per-year `{year}損益表` / `{year}損益表 (可橙)` sheets are the source of truth for the "現價" (current price) column that the app displays on load; `serve.py` reads it directly and the button-triggered fetch writes back into it. Rows are matched by product name/code against `fetch_close.py`'s `STOCK_CODE_MAP`/`ETF_CODES`, scanning each sheet from row 57 onward (below the weekly log) rather than fixed row ranges, so inserting/deleting holding rows in Excel doesn't break the mapping. **Don't edit this file with tools while it's open in Excel** — check for a `~$菡萏咖啡.xlsx` lock file first; writing to it while Excel has it open risks Excel's next save silently clobbering the change.
- `HanDan/prices.json`, `HanDan/embedded_prices.js` — generated output from `fetch_close.py` (not hand-edited).
- `HanDan/.venv` — uv-managed Python 3.14 virtualenv local to this machine; needs `requests`, `openpyxl`, and (only if rebuilding the exe) `pyinstaller` installed.
- `HanDan/菡萏咖啡存摺.exe`, `HanDan/菡萏咖啡存摺.spec` — optional standalone build of `serve.py` via PyInstaller, for running without a Python install. The `.spec` is the reproducible build recipe (rebuild with `pyinstaller 菡萏咖啡存摺.spec`); the `.exe` must stay in `HanDan/` next to the xlsx/html data files (see BASE_DIR note below), not moved elsewhere on its own.

## Commands

Environment setup (Windows + [uv](https://github.com/astral-sh/uv), run from `HanDan/`):

```bash
uv venv
.venv\Scripts\activate
uv pip install requests openpyxl
```

Run the local server (serves the page + price APIs at `http://127.0.0.1:8765`, opens a browser automatically). Day-to-day, double-click `HanDan/開啟存摺.bat` instead — it does the same thing without needing a terminal:

```bash
python serve.py
```

Fetch closing prices standalone and print the replacement JS snippet to stdout:

```bash
python fetch_close.py
```

Fetch closing prices and patch the HTML file directly (backs up the original to `.bak` first):

```bash
python fetch_close.py --html "菡萏咖啡-台股損益存摺.html"
```

There is no lint or test command — there is no test suite or linter config in this repo.

Rebuild the standalone exe after changing `serve.py` or `fetch_close.py` (run from `HanDan/`, needs `pyinstaller` installed per above):

```bash
.venv\Scripts\pyinstaller.exe 菡萏咖啡存摺.spec
copy dist\菡萏咖啡存摺.exe .
rmdir /s /q build dist
```

The `.spec` already includes `--collect-data certifi` (needed for `requests`' TLS verification to work inside the frozen exe) and targets `serve.py` as the entry point; PyInstaller picks up `fetch_close.py` automatically since `serve.py` imports it directly. No hidden-imports flags needed for `openpyxl`/`requests` — the community hooks bundled with PyInstaller handle both.

## Architecture

### fetch_close.py / serve.py data flow

Both scripts share the same closing-price logic (`serve.py` imports `fetch_close` as `fc`):

1. `target_trading_date()` decides which trading day to fetch: before 13:30 → previous trading day, after 13:30 → today, weekends roll back to the nearest weekday (holidays are not detected — the source silently returns the last available trading day instead).
2. TWSE (listed) prices are fetched for that specific date first (`fetch_twse_dated`); if that returns nothing, it falls back to the latest available snapshot (`fetch_twse_latest`, a full CSV dump with no date filter).
3. TPEx (OTC) prices come from a separate open API (`fetch_tpex`), tried against several possible field-name variants since the source has changed schema before.
4. Only symbols in `ETF_CODES` and `STOCK_CODE_MAP` (the hardcoded holdings list) are kept.
5. `fetch_close.py` writes `prices.json` (full price map) and `embedded_prices.js` (a snippet meant to replace the `EMBEDDED_PRICES`/`EMBEDDED_PRICES_DATE` constants in the HTML), and optionally patches the HTML in place via regex (`patch_html`) — this is a textual replace of the `const EMBEDDED_PRICES_DATE = ...; const EMBEDDED_PRICES = {...};` block, not JSON-aware, so that block's exact `const` declaration format in the HTML must be preserved.

`serve.py` splits the same underlying fetch logic across two endpoints, deliberately kept separate so opening the page never hits the network or touches the user's Excel file on its own:

- **`/api/xlsx-prices`** — called automatically whenever the page loads. Reads current prices straight out of `菡萏咖啡.xlsx`'s "現價" column (`read_xlsx_prices()`); no TWSE/TPEx call, so it's fast and works offline. Returns `date`/`targetDate` both set to the xlsx file's mtime (so the front end's staleness check never false-fires here) and a one-line `log` entry naming the source file.
- **`/api/close-prices`** — called only when the user clicks "取得盤後收盤價" (front end passes `?updateXlsx=1`). Does the live `collect_prices()` fetch against TWSE/TPEx as before, then calls `update_xlsx_prices()` to write the results back into `菡萏咖啡.xlsx`'s current-year sheets, so the *next* page load's `/api/xlsx-prices` read picks them up.

`EMBEDDED_PRICES` in the HTML is the final fallback, used only when `serve.py` isn't reachable at all (e.g. opening the HTML file directly instead of via `開啟存摺.bat`/`serve.py`) or when both API calls fail — the front end's `fetchClosePrice()` (菡萏咖啡-台股損益存摺.html) threads a `syncXlsx` flag through to pick the right endpoint and to tell the two failure modes apart in the status message ("not running via serve.py" vs "serve.py is up but reading/writing the xlsx failed").

`update_xlsx_prices()`/`read_xlsx_prices()` in `serve.py` scan each `{year}損益表`/`{year}損益表 (可橙)` sheet from row 57 to the end, matching column-A product names/codes against `STOCK_CODE_MAP`/`ETF_CODES` rather than assuming fixed row ranges — this also naturally picks up the "借券(出借)" (securities-lending) sub-tables further down the sheet, which use the same column layout. Rows whose name isn't in the tracked set (sold/closed positions) are left untouched. Writing to the xlsx via `openpyxl` strips cached values from **every** formula cell in the whole workbook (not just the touched rows) until Excel is opened and recalculates on load — this is expected `openpyxl` behavior, not data loss, but don't be alarmed if a `data_only=True` read shows blanks for unrelated sheets right after a write.

When adding/removing holdings, update `ETF_CODES` / `STOCK_CODE_MAP` in `fetch_close.py` **and** the matching `STOCK_CODE_MAP`/ETF list embedded in the HTML (`菡萏咖啡-台股損益存摺.html:1155`) — they must stay in sync since the front end maps stock names to codes independently for its own display.

`serve.py`'s `BASE_DIR` resolves via `sys.executable`'s parent when `sys.frozen` is set (running as the PyInstaller exe), falling back to `Path(__file__).resolve().parent` otherwise — this is deliberate, not defensive boilerplate: under PyInstaller onefile, `__file__` points into the temp `_MEIPASS` extraction directory, which would make the exe read/write a throwaway copy of the xlsx instead of the real one sitting next to it. If BASE_DIR-based path logic changes, keep this frozen/unfrozen branch working, and re-test by running the actual `.exe` (not just `python serve.py`) after any such change.

**Known TPEx gotchas (fixed 2026-07-31, don't re-diagnose from scratch if they resurface):**
- If TPEx (上櫃/興櫃) requests fail with `SSLError: CERTIFICATE_VERIFY_FAILED ... Missing Subject Key Identifier`, this is not a transient network issue — `www.tpex.org.tw`'s certificate chain lacks the SKI extension, which OpenSSL 3.2+'s default strict mode (`VERIFY_X509_STRICT`) now rejects. `fetch_close.py` works around this with `_RelaxedCertAdapter`, a `requests` Session mounted only on `https://www.tpex.org.tw` that clears the strict flag (hostname/chain validation otherwise unchanged). If this starts failing again, check whether `ssl.VERIFY_X509_STRICT` is still available in the active Python/OpenSSL build.
- The TPEx ESB (興櫃) API's close-price field is `LatestPrice`, not `LastPrice` — a one-character-off name that silently produces zero matches from `_parse_quotes` if missing from `close_keys`, since failures there are swallowed rather than raised. When a whole TPEx dataset comes back "查無報價" for every symbol (not just a few), suspect a field-name mismatch in `_parse_quotes` before assuming the source is down — probe with `python fetch_close.py --probe <code>` to see the raw counts per dataset.

### HTML app structure

Single-page app, no framework. Key pieces (all inline `<script>` starting around line 568):

- **`PORTFOLIO_DATA`** (菡萏咖啡-台股損益存摺.html:569) — one giant minified JSON literal holding the static baseline data: `weekly`, `stocks`, `etf`, `lending`, `overall`, each keyed by account (`main`, `keqiang`) and then by year. This is the historical/baseline ledger and is not meant to be hand-edited casually — treat it as generated/exported data.
- **Two accounts**: `main` (我的戶頭) and `keqiang` (可橙戶頭), threaded through nearly every function and state key as an `acc` parameter/index.
- **`state`** — all live, user-editable overlay data on top of `PORTFOLIO_DATA`: per-account `customWeekly`, `soldOverride`, `fieldOverride` (qty/avg_price/price edits), `confirmedStocks`/`confirmedDelta` (baked-in baseline after user clicks "確認並存檔"), dividend schedules/overrides, sort/filter UI state.
- **Effective-value pattern**: raw baseline rows from `PORTFOLIO_DATA` are never mutated directly. Functions like `effectiveStock`, `effectiveSold`, `effectiveEtf` layer `state` overrides on top of baseline rows at render time, and `computeLiveDelta`/`computeHoldingsDelta` compute the diff between edited and original figures so yearly/weekly summaries (`syncedOverall`, `syncedWeekly`) can reflect live edits without rewriting the baseline data.
- **Storage abstraction** (菡萏咖啡-台股損益存摺.html:589 onward) — `detectStorageBackend()` picks, in order: Claude artifact storage (`window.storage`) → `localStorage` (prefixed `handan:`) → in-memory `Map` (lost on refresh, triggers a warning banner telling the user to export a JSON backup instead). All persistence goes through `storageGet`/`storageSet`/`storageDelete`/`storageList` so the rest of the app is agnostic to which backend is active.
- **Key registry for backups** — every storage key written is registered into `state.knownKeys` (persisted itself under `backup-key-index`) so the "05 備份與還原" export can reliably back up everything; `wellKnownKeyCandidates()` is a fallback list of keys used by older versions of the app, for backups created before the registry existed.
- **UI is tab-based** (菡萏咖啡-台股損益存摺.html:432-436), five tabs switched via `data-tab` buttons: 01 總覽與趨勢 (overview/charts), 02 持股明細 (holdings table, inline cell editing), 03 存股與配息 (ETF/dividend tracking), 04 記一筆 (manual entry), 05 備份與還原 (JSON export/import, and XLSX export/import via SheetJS for holdings).
- Charts (`renderYearlyChart`, `renderWeeklyChart`) are rendered with Chart.js against `syncedOverall`/`syncedWeekly` output, so they always reflect live edits, not just the static baseline.

### Working with this codebase

- All UI strings, comments, and docstrings are in Traditional Chinese (Taiwan usage) — match that when editing this project.
- Since the HTML is one file with no source maps or components, use line-anchored greps (e.g. `EMBEDDED_PRICES`, `PORTFOLIO_DATA`, function names above) to navigate rather than assuming a directory structure.
- `PORTFOLIO_DATA` and `EMBEDDED_PRICES` are both large single-line JSON/JS literals — avoid full-file reads; target them with line offsets or regex search instead.
