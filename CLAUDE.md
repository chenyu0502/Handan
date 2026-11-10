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
- `HanDan/xlsx_cache.json` — 公式快取的備援值，`{"分頁名!座標": 數值}`，由 `serve.py` 自動寫入，gitignored。寫入「現價」時 openpyxl 會清空**整本** xlsx 的公式快取，之後每個公式格都讀成 `None`，直到使用者再用 Excel 開啟並存檔；這個檔記住每次「真的讀到數值」的格子，讀不到時頂替。刪掉會自己長回來。**`菡萏咖啡_手操版.xlsx` 是使用者人工整理的備份，不是資料來源，任何讀取路徑都不得去讀它**（2026-10-21 起；在那之前它曾被當成快取備援，導致網頁顯示的數字取決於使用者最後一次手動存檔的時間）。
- `HanDan/prev_closes.json` — previous-trading-day closes, written by `serve.py` on every price fetch and gitignored. This is the only basis the 今日損益 card has: `fetch_realtime_quotes()` is the sole source that returns a previous close (MIS field `y`), the official daily files carry only the day's own four prices, and the xlsx 現價 column stores a single number. Page load reads the xlsx, so without this file on disk the card would have nothing to compare against and shows 「需按一次取得盤後收盤價」. The stored `priceDate` is surfaced so a stale basis is visible rather than silently wrong.
- `HanDan/vendor/` — `chart.umd.min.js` and `xlsx.core.min.js`, inlined by `build_web.py`. Kept in version control on purpose so the build stays reproducible if those CDN URLs change or disappear. **Use the SheetJS `core` build, never `full`**: the full build's codepage table carries ~51,700 U+FFFD filler characters, which is what originally broke publishing, and core still exposes every SheetJS API the app calls while being 444 KB smaller.
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

Rebuild and deploy the phone-facing page (run from `HanDan/`):

```bash
python build_web.py
```

```bash
cd handan-web && vercel deploy --prod --yes
```

**`vercel deploy` often reports `"status": "error"` on the first run and succeeds on an immediate retry** — seen three times now. Nothing was uploaded when that happens, so always confirm by fetching the live `version.json` and checking it matches the version the build just printed; do not trust the CLI's exit output alone.

**The day-to-day flow needs nothing from Claude at all**: the user opens `開啟存摺.bat`, edits and hits 「取得盤後收盤價」 in the page, and `serve.py` pushes the snapshot to Firestore automatically, so the phone is already current. Only rebuild and redeploy when the *page code itself* changes — a data-only change never needs it.

The `更新手機版` request therefore means "the app changed, redeploy it", not "sync my data".

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
3a. **The official daily files lag the close by roughly an hour** — measured at 13:55 the TWSE dated endpoint still returns 1363 rows dated to the *previous* trading day. When that happens both TWSE and TPEx are stale *and agree with each other*, so comparing sources against each other cannot detect it; an earlier version only cross-compared and therefore silently handed back yesterday's prices to anyone who pressed 「取得盤後收盤價」 right after the close. Both `collect_prices()` (serve.py) and `main()` (fetch_close.py) now check `iso != target_iso` while `before_close` is false, and fill in from `fetch_realtime_quotes()`, the MIS quote engine, which already reflects the final trade price at 13:30. That helper queries each code under **both** `tse_` and `otc_` prefixes (the engine returns `"-"` for the wrong market, which the parser drops) so callers need not know a symbol's market, and it batches at `MIS_BATCH_SIZE` because a single oversized query returns nothing at all. Emerging-stock (興櫃) codes are not covered by MIS and stay on the official file's date — 67 of 68 holdings get same-day prices, and the remaining one is logged.
4. Only symbols in `ETF_CODES` and `STOCK_CODE_MAP` (the hardcoded holdings list) are kept.
5. `fetch_close.py` writes `prices.json` (full price map) and `embedded_prices.js` (a snippet meant to replace the `EMBEDDED_PRICES`/`EMBEDDED_PRICES_DATE` constants in the HTML), and optionally patches the HTML in place via regex (`patch_html`) — this is a textual replace of the `const EMBEDDED_PRICES_DATE = ...; const EMBEDDED_PRICES = {...};` block, not JSON-aware, so that block's exact `const` declaration format in the HTML must be preserved.

`serve.py` splits the same underlying fetch logic across two endpoints, deliberately kept separate so opening the page never hits the network or touches the user's Excel file on its own:

- **`/api/xlsx-prices`** — called automatically whenever the page loads. Reads current prices straight out of `菡萏咖啡.xlsx`'s "現價" column (`read_xlsx_prices()`); no TWSE/TPEx call, so it's fast and works offline. Returns `date`/`targetDate` both set to the xlsx file's mtime (so the front end's staleness check never false-fires here) and a one-line `log` entry naming the source file.
- **`/api/xlsx-weekly`** / **`/api/xlsx-overall`** — also called automatically on page load (and again when "取得盤後收盤價" runs). The first reads the current year's weekly log, the second the 「總績效」 sheet; both are read-only and both are included in the cloud snapshot (`build_snapshot`) so the phone gets the same numbers.
- **`/api/close-prices`** — called only when the user clicks "取得盤後收盤價" (front end passes `?updateXlsx=1`). Does the live `collect_prices()` fetch against TWSE/TPEx as before, then calls `update_xlsx_prices()` to write the results back into `菡萏咖啡.xlsx`'s current-year sheets, so the *next* page load's `/api/xlsx-prices` read picks them up.

`EMBEDDED_PRICES` in the HTML is the final fallback, used only when `serve.py` isn't reachable at all (e.g. opening the HTML file directly instead of via `開啟存摺.bat`/`serve.py`) or when both API calls fail — the front end's `fetchClosePrice()` (菡萏咖啡-台股損益存摺.html) threads a `syncXlsx` flag through to pick the right endpoint and to tell the two failure modes apart in the status message ("not running via serve.py" vs "serve.py is up but reading/writing the xlsx failed").

**所有讀 xlsx 數值的路徑都走 `xlsx_cell(ws_cached, ws_raw, coord)`**（`read_xlsx_weekly`、`read_xlsx_overall`、`read_xlsx_balance`、`_find_pending_dividend`、`_read_balance_and_pending`、`_read_existing_stock_dividend`）。順序是：

1. 主檔公式快取（讀到就順手寫進備援表）
2. 公式版的字面數值
3. 自己算公式：`_safe_eval_arith` 處理純算式（L 欄的 `=0.699*1000-10`），`_resolve_formula` 再處理含儲存格參照的（`=$B$1`、`=H2-N2`、`='2025損益表'!B1`），參照會遞迴追下去
4. `xlsx_cache.json` 備援表

公式自己算排在備援表前面，因為它反映「現在的公式」，備援表只反映「上一次讀到的結果」。新增讀取路徑時沿用這支，不要自己讀 `.value`。

**第 3 步的參照解析不是可有可無的最佳化。** 損益表整欄 B（股本）是 `=$B$1`，B1 是 `=H2-N2`，H2 是 `='2025損益表'!B1`——寫入現價清掉快取後這些全部讀成 None，少了這一層，我的戶頭 2026 的週記帳會從 41 列掉到 21 列（六月之後整批消失），當週損益變成跟五月比、圖表也斷在五月。`_resolve_formula` 碰到函式呼叫或範圍（`=SUM(G2:R2)`）就放棄，交給備援表；循環參照用 `seen` 集合擋住（不是遞迴層數——結果要跟「從哪裡開始追」無關，上層才能安心快取）。解析結果（含失敗）快取在公式版 workbook 上，否則同一條鏈會被重走上千次（實測 5.7 秒 → 0.8 秒）。

**「有沒有函式呼叫」要先剝掉引號裡的分頁名才能判斷 (fixed 2026-11-09)。** `_UNSUPPORTED` 是 `[0-9A-Za-z_一-鿿]\s*\(|:`，而分頁名自己就帶括號——`'2025損益表 (可橙)'!E2` 裡的 `表 (` 會命中，於是**可橙每一條跨年度參照都解析不出來**。後果不是少一個最佳化：可橙的 `B2`（帳戶餘額）整條鏈都指向前一年度分頁，原本只能靠備援表活著，備援表一冷掉帳戶餘額就空白。現在走 `_unsupported()`，它先用 `_QUOTED_SHEET` 把 `'...'` 換掉再判斷；`=SUM(G18:R18)` 照樣會被擋（括號與冒號都還在）。

**`SUM(單一連續範圍)` 現在解得出來 (2026-11-09)。** `_SUM_CALL` 在一般的參照代入**之前**先把 `SUM(D58:D74)`、`SUM('2025損益表'!A1:A9)` 換成算好的總和，之後才檢查 `_unsupported()`——所以剩下的括號／冒號就是真的不支援的（巢狀、多參數 `SUM(A1:A9, B1)`、其他函式）。範圍裡的空白格算 0，但只要有一格「有內容卻算不出來」就整條放棄，不要吐一個少算幾格的總和；`SUM_MAX_CELLS` 擋住大到離譜的範圍。**為什麼需要**：我的戶頭 `B2` 鏈上掛著 `-M100`，而 `M100 = =SUM(M58:M74)`，少了這一層 B2 永遠只能讀備援表的舊值——登記出借收入之後帳戶餘額不會動，就是這樣來的。改動後 weekly／overall／balance／cash／prices 五組讀數與改前**逐字元相同**（`diff` 0 行），耗時 0.86s vs 0.95s；差別只在這些值現在是即時算出來的，不再依賴備援表的新鮮度。

**參照到真正空白的格子要當成 0，不是失敗 (fixed 2026-11-09)。** Excel 的算式裡空白格就是 0：`E2 = =$B$2-$J$2+J48` 的 `J48`（本週買賣）在還沒交易的週次是空的，可橙 2025 的 `=9693-K2+J2` 裡 `K2` 一直空著。原本 `sub()` 一碰到 None 就 `failed = True`，一個空格就讓整條鏈斷掉、`E2` 永遠讀不到。現在只有「那一格在公式版有內容卻算不出來」才算失敗；公式版的 `.value` 是 None（既無公式也無字面值）就代入 0。改動後 weekly/overall/balance/prices 四組讀數與改前**逐格相同**，只有兩格從備援表的 float 變成公式算出的 int（`57008.0`→`57008`、`2156675.0`→`2156675`）。

**寫入前一定要先養備援表。** `=SUM(...)` 這種公式 `_resolve_formula` 追不下去，快取一清就真的沒有了。所有「載入→改→存檔」都走 `_open_for_write()`，它會先呼叫 `warm_memo_before_write()` 把三支讀取跑一遍。**不要直接 `openpyxl.load_workbook(xlsx_path, data_only=False)` 然後存檔**，那次寫入會把還沒記錄的快取永久清掉。

**「總績效」整張表都是追不下去的那一類**，每一格最終都指向 `{year}損益表!C55` 這種格子，而那些 C55 是 `=$F$138+$E$2+$L$136-$C$2`，F138 又是 `=SUM(...)`。所以 20 格（第 3/4 列與第 9/10 列的 C~H）只能從 Excel 重算後的快取取得，備援表一旦冷掉就整張表空白，前端會退回 `PORTFOLIO_DATA.overall`，可橙 2026 的 −1,460,000 提領就不見了。2026-10-22 已用 10-20 從主檔快取讀到的值把這 20 格補進備援表，之後由 `_open_for_write()` 的預先養表維持。

**陷阱：「當年結算」不等於該年度週記帳最後一筆有值的列。** 總績效指的是 C54/C55 那一格（整年最後一列），那格是公式；往回找最後一個「字面數值」會拿到更早的一週，數字看起來合理但全錯（實測 2022 會拿到 468,583 而不是 462,468，2024 拿到 711,108 而不是 695,232）。不要用這個方式補算。

`read_books()` 以 mtime+size 當鍵共用唯讀的活頁簿對（快取版＋公式版），三支讀取函式各自載入要三秒以上，而開啟頁面時它們是連續呼叫的。寫入用的活頁簿不走這裡。

實測：把整本快取清空後（且目錄下沒有手操版），weekly/overall/balance 三組結果與清空前的最大相對誤差 2e-14。

`update_xlsx_prices()`/`read_xlsx_prices()` in `serve.py` scan each `{year}損益表`/`{year}損益表 (可橙)` sheet from row 57 to the end, matching column-A product names/codes against `STOCK_CODE_MAP`/`ETF_CODES` rather than assuming fixed row ranges — this also naturally picks up the "借券(出借)" (securities-lending) sub-tables further down the sheet, which use the same column layout. Rows whose name isn't in the tracked set (sold/closed positions) are left untouched. Writing to the xlsx via `openpyxl` strips cached values from **every** formula cell in the whole workbook (not just the touched rows) until Excel is opened and recalculates on load — this is expected `openpyxl` behavior, not data loss, but don't be alarmed if a `data_only=True` read shows blanks for unrelated sheets right after a write.

When adding/removing holdings, update `ETF_CODES` / `STOCK_CODE_MAP` in `fetch_close.py` **and** the matching `STOCK_CODE_MAP`/ETF list embedded in the HTML (`菡萏咖啡-台股損益存摺.html:1155`) — they must stay in sync since the front end maps stock names to codes independently for its own display.

**「加入存摺」會同步寫一筆進損益表的 D／E 欄**（`append_tx_to_xlsx`，`/api/append-tx`）。買進加進 D 欄、賣出加進 E 欄，列是 `_weekly_row_for_date()` 找的「第一個 A 欄記帳日 >= 交易日」那一列（A 欄是週五，年度最後一列 12/31 不是週五，用同一條規則也會落對）。寫法沿用使用者自己的記法：那兩格長得像 `=23119+23019`，一筆交易就是一個加項，空格時寫成 `=金額` 而不是純數字；同一格的儲存格註解附加一行「商品 單價*股數」。**公式的加項是價金（買進含手續費、賣出已扣費用與交易稅），註解的乘積是未扣費用的成交條件，兩者本來就差一個費用**——使用者手寫的那幾列也是這樣（D47 `=23119+23019` 配註解 `台積電 2310*10`／`2300*10`），不要「修正」成一致。

這條路徑不碰 B2：E2（可用現金）是 `=$B$2+J47`、J47 是 `=E47-D47`，所以寫進 D／E 之後 Excel 的可用現金自己會動；`importTxByDate()`（「匯入」鈕）只改瀏覽器裡的帳戶餘額、不寫 xlsx，兩邊不會重複計算。寫入失敗（最常見是 Excel 正開著那個檔）**不回頭刪掉已經記好的交易**，而是排進 `state.pendingXlsxTx`（storage key `pending-xlsx-tx`），04 記一筆下方出現「補寫 N 筆到 Excel」讓使用者關掉 Excel 後一次補完；補寫時第一筆失敗就停手並把剩下的原封不動留在佇列，不然每筆都要重開一次活頁簿（一次好幾秒）。刪除交易紀錄一樣不回沖 Excel，與配息、持股明細的處理一致。

**同一次寫入還會維護損益表下半部的持股表**（`_sync_holding_row`）。買進：單一相符列就累加庫存數量、均價重算成 `(原成本金額＋價金)/新總股數`；沒有相符列就複製 `Sample` 範本列。賣出：扣完有剩就相減並重算 `(原成本金額−價金)/剩餘股數`，歸零就先把整列複製到 `Sample` 下面留底再清空原列。庫存數量原本是公式的（ETF 整欄 B 長得像 `=3418+140+135`）照 `_append_term_to_formula` 的風格接一個加減項，不要換成單一數字。讀舊的 B／C／D 必須走 `xlsx_cell()`（借 `read_books()` 的唯讀活頁簿對，`_open_for_write()` 已經載過所以是快取命中）——寫過的檔案公式快取是空的，直接讀 `.value` 會拿到 `None`，連續兩筆同商品的第二筆就會失敗。

**絕對不要用 `insert_rows()`／`delete_rows()` 改這張分頁。** openpyxl 搬動儲存格時**不會**修正公式，而這裡滿是跨列的絕對參照（`=$F$138+$E$2+$L$136-$C$2`、`=SUM(D95:D111)+SUM(D121:D134)`），連「總績效」都直接指名 `='2026損益表'!$C$55`——插一列或刪一列會讓整本活頁簿靜默算錯且不報錯。所以新增與出清都只動 `Sample`（整張分頁最後一列）以下的空白區（`_staging_row`），那裡不在任何 `=SUM(...)` 範圍內；「整列刪除」改成 `_clear_holding_row()` 清空 A~L（成本與市值歸零，對總計的效果與刪列相同），空殼留給使用者在 Excel 裡刪。實測一筆出清只改動 17 格，其他分頁與其他列的公式分毫未動。

**J 欄與 K 欄是整個區塊合併成一格的**（`J121:J128`、`K95:K111`、`J131:J135`，可橙是 `J82:J89`、`K94:K100`…）。openpyxl 把合併區域裡非左上角的格子做成 `MergedCell`，`.value` 是唯讀的，寫下去就是 `AttributeError: 'MergedCell' object attribute 'value' is read-only`（2026-11-06 使用者賣國建時踩到）。所有寫入都走 `_writable_cell()`，它對 `MergedCell` 回 `None`，呼叫端跳過——那一格屬於整個區塊，不屬於這一列。

**那次的真正傷害不是例外本身，是「半套結果被存檔」**：留底列已經複製完成、原列只清掉 A~I 就拋出，而例外被 catch 之後照樣 `wb.save()`，L124 就留在那裡。所以 `append_tx_to_xlsx` 現在**先做持股表、後寫 D／E 欄**，持股表拋例外就把整本活頁簿丟掉重載（`_open_for_write()` 再來一次）再只寫 D／E。順序不要再換回去。

`_holding_rows()` 要排除三種列：表頭（B 欄是**非公式**的字串，`=3418+140` 這種公式字串才是持股）、A 欄空白的小計列，以及「借券(出借)」子表（0056 同時在第 95 與第 114 列，加到出借那一列是錯的；借券區從 `借券(出借)` 標記起算到下一個 A 欄空白為止）。**同名多列時刻意不自動合併**——使用者把不同批次分開記（品安 4 列、可橙 00937B 3 列），猜錯一列就改到真實部位；買進改成在 `Sample` 下面開一列，賣出直接跳過並回報。`holding_code()` 是這裡與 `read_xlsx_prices`／`update_xlsx_prices` 共用的名稱→代號換算，額外認得「中文 代號」的寫法（新增列就是這樣填的），否則那一列永遠抓不到現價。

持股表出錯**不中斷** D／E 欄那筆（已經寫好了），只在回傳的 `holding` 欄位回報，前端 `holdingNote()` 翻成一句話；這種情況也**不**排進 `pendingXlsxTx`，補寫會讓金額記成兩筆。

**對照表是單向的（名稱 → 代號），顯示時要反查。** 使用者在 04 記一筆 填代號（`2330`）時，`mergeTxIntoHoldings()` 建出來的列名稱就是代號本身，而 `codeForName()` 對代號是原樣回傳，所以 `nameWithCode()` 會變成「代號＝名稱」而只顯示一串數字。`nameByCode()` 用 `STOCK_CODE_MAP` 反查補上中文（同代號多名稱時取表裡**先**出現的那個——2883 的「凱基金」在舊名「開發金」前面），`canonicalName()` 則在建列與 `confirmHoldingsForAccount()` 烘焙基準值時把名稱換成中文。ETF 本來就以代號為名、不在這張表裡，反查不到就維持原樣。反查表延後到第一次用到才建：雲端版的 `STOCK_CODE_MAP` 是登入後才注入的。

`mergeTxIntoHoldings()` (04 記一筆 → 持股明細) can create a holding row for a product in neither map, so a position may now exist that the price fetch never covers. The UI says so when it happens ("取盤後收盤價時抓不到現價"), but the maps still need updating by hand before prices start flowing. Note how new rows are added: appended to `state.confirmedStocks[acc]` with **qty 0**, with the real position written to `fieldOverride`. Writing the figures straight into the baseline would make `computeLiveDelta` see a zero diff, and 總覽's 庫存市值 would silently not move.

`serve.py`'s `BASE_DIR` resolves via `sys.executable`'s parent when `sys.frozen` is set (running as the PyInstaller exe), falling back to `Path(__file__).resolve().parent` otherwise — this is deliberate, not defensive boilerplate: under PyInstaller onefile, `__file__` points into the temp `_MEIPASS` extraction directory, which would make the exe read/write a throwaway copy of the xlsx instead of the real one sitting next to it. If BASE_DIR-based path logic changes, keep this frozen/unfrozen branch working, and re-test by running the actual `.exe` (not just `python serve.py`) after any such change.

**即時報價（MIS）永遠不是收盤價，官方每日檔才是 (fixed 2026-11-06).** 13:25~13:30 的收盤集合競價不會進到這個引擎，`trade.z` 停在收盤前最後一筆連續撮合；而 14:00 盤後定價交易一開盤，`z` 欄會**整批**變成 `"-"`，整份清單就只剩 `trade.z` 可用。實測 2026-11-06 14:17：台積電 `t`／`trade.t` 都停在 13:19，`trade.z` 是 2140，官方收盤價 2135；58 檔上市持股有 **46 檔**與官方收盤價不符（最大 0.94%，2330 差 5 元）。原本的快路徑把即時報價當權威、官方每日檔只用來補缺漏，所以 14:00 之後按「取得盤後收盤價」寫進 xlsx 的幾乎整批是錯的（使用者 2026-11-06 14:11 按下那次就是）。

**抓價的 log 只寫結論，不寫過程 (2026-11-10).** 原本七行逐一交代每支來源拿到幾檔，使用者反映太長；而且「⚠ 記得稍後再按一次」那句被淹在中間，實際後果是 2026-11-09 的 00940 用盤中價 12.09 留了一整晚（官方收盤價 12.03，20,000 股差 1,200，隔天「今日損益」127,245 因此比實際多 1,200，當週損益 171,946 才是對的）。現在 `overlay()` 與即時報價成功都不寫 log，結論一律由 `_finalize_prices()` 產出：沒抓到就 `⚠ 追蹤清單 N 檔只取得 M 檔`，全是官方價就 `{日期} 收盤價：追蹤清單 N 檔全數為官方公布價`，有暫用值就寫成「官方 N 檔、暫用盤中最後成交價 M 檔」再加兩行（代號清單＋「約 15:30 之後再按一次」）。`pending` 由快路徑算好傳進去。只有「抓取失敗」「興櫃暫用該日價」「當日無成交」這三種需要使用者知道的才另外留一行。前端 `fetchClosePrice()` 把這些**逐行**印（不要用 `/` 串成一條，那句提醒會被淹掉），開頭 `⚠` 的標成警示色；「已收盤（13:30 後）…」那句拿掉了，日期已經在結論那行，只有「尚未收盤取前一交易日」才還需要說明。

`collect_prices()` 的快路徑現在是「即時報價當底、官方檔逐一覆蓋」：三支全市場檔（上市指定日／上櫃／興櫃）只要 `roc_to_iso(raw_date) == target_iso` 就蓋掉即時報價，不是只填缺漏。即時報價剩下兩個不可取代的職責：**官方檔還沒發布那段空窗的頂替**，以及**前一交易日收盤價**（`y` 欄，官方檔沒有，今日損益算不出來）。興櫃刻意排最後且允許「未更新也先用」——即時報價引擎不提供興櫃（tse/otc/emg/rot 四種前綴皆實測無成交價），整批丟掉會讓那幾檔沒有價。整趟實測 1~2 秒。`fetch_close.py` 的 `main()` 本來就是官方優先、只在 `iso_date != target_iso` 時用即時報價補，不必跟著改。

**全市場檔補不到的，用「單檔單月」端點補 (2026-11-06).** 上櫃的全市場 OpenAPI（`tpex_mainboard_daily_close_quotes`）發布得很晚——實測 15:18 還停在前一交易日，所以光靠全市場檔那 10 檔上櫃持股一整天都拿不到正式收盤價。同一份官方資料另外有「一檔一個月」的端點，**發布早得多**，15:18 已經給得出當日收盤價：

- 上市：`https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?stockNo=<代號>&date=<YYYYMM>01&response=json`，欄位 `日期`／`收盤價`，`fields`/`data` 在頂層
- 上櫃：`https://www.tpex.org.tw/www/zh-tw/afterTrading/tradingStock?code=<代號>&date=<YYYY/MM>/01&id=&response=json`，欄位 `日 期`／`收盤`（有全形空格），資料包在 `tables` 陣列裡

`fetch_official_closes(markets, d)`（`fetch_close.py`）就是這一層，`collect_prices()` 只對「全市場檔沒確認到、且市場別不是興櫃」的代號呼叫它，正常情況一檔都不用打。**市場別是用全市場檔的「成員資格」判斷的**（`market_of`）——那份檔案即使日期還是前一交易日，涵蓋哪些代號仍然是對的，所以不必為了分辨上市／上櫃再查一次。日期格式是 `115/11/06`，`roc_to_iso` 只吃純數字，要先濾掉斜線；無成交的日子四價是 `--`，`_to_price` 會回 None 而跳過整列。

**證交所那一側會擋連續查詢**：快速送出 68 檔時有 16 檔收到 HTTP 307，導到一頁「因為安全性考量，您所執行的頁面無法呈現」（不是 JSON，所以會在 `fetch_stock_month` 裡拋 `JSONDecodeError` 而被 `one()` 吞掉）。因此 `MONTH_WORKERS` 按市場分開：證交所 2、櫃買 6（櫃買實測 10 檔 0.12 秒，沒被擋過）。這一層是補強不是主力，被擋的代號留在 log 的「⚠ N 檔官方收盤價尚未發布」裡，晚一點再按一次就會修正。交叉驗證：52 檔同時走單檔端點與全市場檔，兩邊數字零差異。

**MIS 的 `z` 欄不是「今天最後成交價」(fixed 2026-10-21).** 它是*當前撮合區間*的成交價，所以冷門商品只要收盤那一盤沒成交，`z` 就是 `"-"`，即使當天明明有成交量。真正的最後一筆成交留在 `trade.z`。實測 00944 在 2026-10-21 成交 94 張、最後一筆 09:56:30 成交 22.34（與證交所官方收盤價一致），`z` 卻是 `"-"`，整檔被判成「當日無成交」而用昨收 21.97 頂替，10,000 股就差了 3,700。`fetch_realtime_quotes` 因此在 `z` 無效時改用 `trade.z`，但**只在 `v`（當日累計成交量）大於 0 時**——真的整天沒成交時 `trade` 裡留的是前一個交易日的舊價。這一層仍然需要（官方檔空窗期的唯一價格來源），但要記得它給的是連續撮合的最後一筆，**不等於收盤價**，見上一段。

**Known TPEx gotchas (fixed 2026-07-31, don't re-diagnose from scratch if they resurface):**
- If TPEx (上櫃/興櫃) requests fail with `SSLError: CERTIFICATE_VERIFY_FAILED ... Missing Subject Key Identifier`, this is not a transient network issue — `www.tpex.org.tw`'s certificate chain lacks the SKI extension, which OpenSSL 3.2+'s default strict mode (`VERIFY_X509_STRICT`) now rejects. `fetch_close.py` works around this with `_RelaxedCertAdapter`, a `requests` Session mounted only on `https://www.tpex.org.tw` that clears the strict flag (hostname/chain validation otherwise unchanged). If this starts failing again, check whether `ssl.VERIFY_X509_STRICT` is still available in the active Python/OpenSSL build.
- The TPEx ESB (興櫃) API's close-price field is `LatestPrice`, not `LastPrice` — a one-character-off name that silently produces zero matches from `_parse_quotes` if missing from `close_keys`, since failures there are swallowed rather than raised. When a whole TPEx dataset comes back "查無報價" for every symbol (not just a few), suspect a field-name mismatch in `_parse_quotes` before assuming the source is down — probe with `python fetch_close.py --probe <code>` to see the raw counts per dataset.

### HTML app structure

Single-page app, no framework. Key pieces (all inline `<script>` starting around line 568):

- **`PORTFOLIO_DATA`** (菡萏咖啡-台股損益存摺.html:569) — one giant minified JSON literal holding the static baseline data: `weekly`, `stocks`, `etf`, `lending`, `overall`, each keyed by account (`main`, `keqiang`) and then by year. This is the historical/baseline ledger and is not meant to be hand-edited casually — treat it as generated/exported data.
- **Two accounts**: `main` (我的戶頭) and `keqiang` (可橙戶頭), threaded through nearly every function and state key as an `acc` parameter/index.
- **`state`** — all live, user-editable overlay data on top of `PORTFOLIO_DATA`: per-account `customWeekly`, `soldOverride`, `fieldOverride` (qty/avg_price/price edits), `confirmedStocks`/`confirmedDelta` (baked-in baseline after user clicks "確認並存檔"), dividend schedules/overrides, sort/filter UI state.
- **Sold rows are not kept.** Two things mark a holding for removal: `mergeTxIntoHoldings()` selling it to zero from 04 記一筆, and 02 持股明細 每列「狀態」欄下拉選單（`select.row-action`，收了暫停／還原／刪除三個動作，第一項顯示目前狀態）的「刪除這一列」。後者是給「記錯了」「測試留下的列」用的——刪掉 04 記一筆的交易紀錄**不會**把持股明細改回去（那筆已經併進庫存了），在這顆按鈕之前只能靠「記一筆賣出歸零」間接清掉，但那會連帶寫進 Excel 的 D／E 欄並調整帳戶餘額，對一列本來就不該存在的資料是錯的。兩者都只是 `state.soldOverride[acc][id] = true`，所以下面那套（立刻隱藏、成本市值歸零、下次 `confirmHoldingsForAccount()` 從基準值移除並先下載備份）兩種來源共用；實測刪掉一列 10 股 @1234 的測試列，costDelta 與 valueDelta 各減 12,340，烘進基準值後兩者都不再變動。There is no "show sold" filter. A sold row is hidden at once and dropped from the baseline on the next `confirmHoldingsForAccount()`; that function downloads a separate `…-刪除已賣出前.json` backup *before* baking whenever it will drop rows, because its usual backup runs after the bake and would no longer contain them. Dropping them is delta-neutral: the round's live delta is folded into `confirmedDelta` first, and sold rows contribute zero cost/value in `computeLiveDelta` either way.
- **Dividend data is per-year.** 03 存股與配息 has a year dropdown. Storage keeps one key per type per account (`div-override:{acc}` etc.) holding `{year: {name: {month: ...}}}`. In memory the four `*All` fields hold the full object while `state.divOverride[acc]` and friends are **references into the selected year's object**, so the ~40 existing `state.divOverride[acc][name][m]` call sites needed no change and writes land in the right year automatically. `bindDivYear()` maintains those references — never assign `state.divOverride[acc] = {...}` wholesale or the link breaks. `normalizeDivByYear()` migrates the old un-dated shape, distinguishing a year key from an ETF name like `0050` by requiring the value to fall in 2000–2100. `PORTFOLIO_DATA.etf[].monthly_dividends` belongs to `DIV_BASE_YEAR` only, so other years come up blank. `pendingDivImports()` deliberately scans the `*All` objects across every year — an overdue import does not stop mattering because the view switched years. The old 「清空所有 ETF 配息設定」 button existed to reset at year end and is gone.
- **帳戶餘額（B2）不等於可用現金（E2）(fixed 2026-11-09).** 可橙的損益表有一欄「基金存款」（J1，J2 是累計額、M1/M2 是累計扣款），那筆錢鎖在定期定額基金裡不能買股票，所以 `B2` 含它、`E2 = $B$2-$J$2+J48` 把它扣掉——2026-11-09 實測 B2 57,008、J2 32,543、E2 24,465。**週記帳的「即時庫存」C 欄用的是 E2**（`C47 = =$F$80+$E$2+$L$78-$C$2`），而 `effectiveCash()` 原本讀 B2，所以可橙的「當週損益」每週固定多算 32,543，而且不會自己消失（我的戶頭 `E2 = =$B$2+J48`，沒有 J2 這一項，兩者相等，所以看不出來）。`read_xlsx_balance()` 因此多回一個 `cash`，值是 **`E2 − J{本週列}`**：`E2` 已經含 Excel 那側本週的買賣淨額（`J48`），而前端另外用 `currentWeekTxNet()` 算同一件事，直接給 `E2` 會把本週交易算兩次；扣掉那一格就是「本週還沒買賣時的可用現金」，本週交易仍由前端加上去。`effectiveCash()` 改用它，並**刻意不理帳戶餘額的手動覆寫值**——那是 B2 口徑的數字，基金存款的差額不會因為使用者改了餘額而改變，硬套只會把那筆錢加回來；要覆寫可用現金請用它自己的欄位。寫完 B2 之後（存簿資金輸入）要重新 `loadAccountBalanceBaseline()`，否則 `accountCashBaseline` 會落後那一筆。`build_snapshot()` 直接收 `read_xlsx_balance()` 的回傳、`build_web.py` 攔 `/api/xlsx-balance` 回 `s.balance`，所以手機版自動跟著拿到 `cash`，兩邊都不必改。
- **第 2 列的累加欄位：「基金扣款管理」與「出借收入管理」(2026-11-09).** 03 存股與配息的兩顆按鈕 → `openLedgerModal(kind, btn)` → POST `/api/ledger-append` → `append_ledger_entry()`，兩者共用同一套流程，差別只在 `LEDGER_KINDS` 這張表（`fund` 可橙「基金扣款」、`lending` 我的戶頭「出借收入」）與前端的 `LEDGER_UI` 文案。**欄位一律用第 1 列的標籤找（`_label_cell()`），不可以寫死欄名**：兩者剛好都在 `M2` 但意思完全不同，寫錯欄會把另一筆資料整格污染掉；找不到標籤就拒絕，按鈕也由 `renderLedgerBtns()` 隱藏（手機版一律不顯示，那邊是唯讀的）。寫入只做兩件事：`_append_delta_to_cell()` 在那一格公式後接 `+金額`，以及 `_append_ledger_note()` 在同一格註解記日期，格式由 `note` 決定——`month-dates`（基金扣款）同月接在最後一行後面（`11/3` → `11/3、11/9`）、跨月才換行，沿用使用者自己的 `10/3、10/13、10/27` 記法（2026 的 M2 有 31 個 `+5000`，正好等於註解裡 10 個整月 ×3 ＋ 11/3，兩邊必須保持對得上）；`line-amount`（出借收入）一筆一行帶金額（`11/9 +196`），因為金額不固定，只記日期會看不出對應公式裡的哪一項。**絕對不要順手再去改帳戶餘額**：兩個帳戶的 `B2` 公式本身就引用那一格——可橙 `B2 = …+J2+…` 而 `J2 = 前一年度!J2 - M2 + 本年存入`，所以扣款一加，J2 與 B2 自己跟著減、`E2 = B2 - J2` 不動（沙箱實測 M2 +5,000 → J2 32,543→27,543、B2 57,008→52,008、E2 24,465 不變）；我的戶頭 `B2 = …+M2+…`，出借收入一加，B2 與 E2 都跟著加（實測 +399.5 → B2 10,098.75→10,498.25）。自己再加一次就是重複計算。金額限正數，負的沖銷要在 Excel 做——`month-dates` 那種只記日期，看不出正負。寫完前端重新 `loadAccountBalanceBaseline()` 而不是自己推算，讓 Excel 的公式當準。`_append_delta_to_cell()` 的加項用 `_fmt_num()`，否則會寫成 `+5000.0`（delta 一路都是 float），跟使用者手寫的 `+5000` 不一致。
- **損益類別的存簿資金輸入在網頁上沒有任何清單。** 視窗下方那張表只列 `capitalFlows`，而 `capitalFlows` 只在 `if (capital)` 時才寫（見下一條），所以折讓、利息這類「損益」項目唯一的紀錄是 xlsx `B2` 的公式加項與**同一格的儲存格註解**（`11/6 +1636 折讓入賬` 這種格式）。使用者 2026-11-09 就是在網頁上找不到一筆 11/6 的 1,636 折讓入帳——它在 Excel 裡，兩個地方都有。要回答「那筆錢哪來的」就去讀 `B2` 的註解。
- **Capital flows are not P&L.** 存簿資金輸入 (03 存股與配息) tags each entry as either 損益 (interest/fees) or 資金進出 (deposit/withdrawal). A withdrawal lowers 可用現金 and therefore 即時庫存, so without a matching cut to the 股本 baseline `renderYearlyMetrics()` books it as a 300k loss — that was a real bug. 資金進出 entries are stored per account in `state.capitalFlows` (`capital-flows:{acc}`) as `{date, amount, note}`, and `capitalFlowSince(acc, lastWeek.date)` adjusts `cost` in one place; 當週損益 and 當年損益 both derive from `currentPnl`, so that single adjustment fixes both, and 績效's denominator uses the adjusted cost too. The filter is strictly `date > lastWeek.date`, which makes the adjustment retire itself once the user writes the next weekly row in the xlsx with the new 股本 — no double counting. The modal's 「只補登，不改帳戶餘額」 records a flow without touching the balance or the xlsx, for money that moved before this existed. 兩種類別寫進帳戶餘額與 xlsx B2 的數字完全一樣，**唯一的差別就是有沒有多記一筆 `capitalFlows`**，所以「類別選錯」的修正方式就是把那一筆 flow 刪掉（餘額與 Excel 都不必動）。視窗下方因此列出該帳戶所有的 `capitalFlows`，每筆可單獨刪除，`capitalFlowBaselineDate()` 之前的會淡化標為「已退場」。
- **01 總覽兩個績效的分母是使用者自填的，不是推算的.** 年度選單下面有一個「當年成本」輸入框 (`yearCostBaseInput`)＝年初庫存，每個帳戶、每個年度各一個值，存在 `year-cost-base:{acc}` → `{year: number}`，`yearCostBaseOf()` 讀它。它同時是「當年績效」(`renderYearlyMetrics`) 與「累計績效／週績效」(`combinedWeekly` 末段重算 `cum_pct`/`wow_pct`，`syncedWeekly` 同步跟進) 的分母。**只有分母換掉**：當年損益仍是 Σ`week_pnl` ＋ 即時 delta 並受 `capitalFlowSince` 調整，`cum_pct` 的分子仍是各列自己的 `inventory − cost`。分母刻意**不**跟著資金進出動——xlsx 的 B 欄會隨年中提領縮小（可橙 2026 從 4,122,506 掉到 2,222,506），分母跟著縮會讓累計績效線憑空跳一段（87.65% vs 正確的 47.25%）。沒填的年度完全退回舊行為，輸入框的 placeholder 會顯示那個退回值；清空欄位等於刪掉該年度的設定，不是存成 0。`YEAR_COST_BASE_SEED` 只在 storage 完全沒有這個 key 時寫入一次；它含兩戶的本金金額，所以列在 `build_web.py` 的清除名單裡。serve.py 的 `read_xlsx_weekly()` 仍照 Excel I 欄原規則算 `cum_pct`，重算一律在前端做——那個值只有瀏覽器儲存裡才有。持股明細的 `perf` 與 ETF 的 `roi` 是另一回事（個別商品的賺賠 ÷ 買進成本），跟這裡無關。
- **「歷年庫存績效走勢」直接以 xlsx 的「總績效」分頁為準.** `read_xlsx_overall()`（serve.py）讀那張表：我的戶頭在第 1~6 列、可橙在第 7~12 列，年度橫排在 C~L 欄；回傳每年的 `inventory`（年初庫存）、`settlement`（當年結算）、`pnl`、`performance`、`cashAdjust`。關鍵是**年初庫存已含當年的現金加碼／提領**（第 1 列／第 7 列），`PORTFOLIO_DATA.overall` 那份靜態快照沒有這一層，可橙 2026 會少掉 −1,460,000、績效從 24.15% 變成 23.26%。前端 `loadXlsxOverall()` → `state.xlsxOverall`，`syncedOverall()` 逐年以 xlsx 覆蓋靜態資料，xlsx 沒有的年度才留靜態值。**今年的「當年結算」不讀 Excel**，改用 `realtimeInventory(acc)`＝庫存市值＋可用現金＋未入帳配息，也就是 C55 那條公式（`=$F$138+$E$2+$L$136-$C$2`）在算的同一件事：寫入現價會清掉 C55 的快取，要等使用者在 Excel 重算存檔才更新，這張卡片因此會停在上一次存檔的數字（實測差了 57,204）。實測即時值與 Excel 剛存檔時的 C55 分毫不差。那一列會帶 `live: true`，tooltip 也會標明。取不到即時值（例如可用現金未設定）才退回 Excel 的結算值，此時若來源是 xlsx 就不再疊加 `computeHoldingsDelta`（理由同 `syncedWeekly`：口徑不同）。這連帶讓頁首三個 chip 與「歷年累計損益」也改用 Excel 的數字，後者現在會等於「總績效」的 A5／A11。快取被清空時靠 `BACKUP_XLSX_NAME` 補；`num()` 先確認公式版那格不是空白才去載備份檔，否則光是 2027~2030 的空欄位就會白白多花 0.4 秒。年初庫存真的讀不到時只對**最新年度**推算（前一年結算＋現金調整，正好是那一格的公式），更早的年度公式長相不一致（2022 讀的是當年第一週庫存），寧可讓那年消失也不要湊出看似合理的錯值。
- **Effective-value pattern**: raw baseline rows from `PORTFOLIO_DATA` are never mutated directly. Functions like `effectiveStock`, `effectiveSold`, `effectiveEtf` layer `state` overrides on top of baseline rows at render time, and `computeLiveDelta`/`computeHoldingsDelta` compute the diff between edited and original figures so yearly/weekly summaries (`syncedOverall`, `syncedWeekly`) can reflect live edits without rewriting the baseline data.
- **Storage abstraction** (菡萏咖啡-台股損益存摺.html:589 onward) — `detectStorageBackend()` picks, in order: Claude artifact storage (`window.storage`) → `localStorage` (prefixed `handan:`) → in-memory `Map` (lost on refresh, triggers a warning banner telling the user to export a JSON backup instead). All persistence goes through `storageGet`/`storageSet`/`storageDelete`/`storageList` so the rest of the app is agnostic to which backend is active.
- **Key registry for backups** — every storage key written is registered into `state.knownKeys` (persisted itself under `backup-key-index`) so the "05 備份與還原" export can reliably back up everything; `wellKnownKeyCandidates()` is a fallback list of keys used by older versions of the app, for backups created before the registry existed.
- **UI is tab-based** (菡萏咖啡-台股損益存摺.html:432-436), five tabs switched via `data-tab` buttons: 01 總覽與趨勢 (overview/charts), 02 持股明細 (holdings table, inline cell editing), 03 存股與配息 (ETF/dividend tracking), 04 記一筆 (manual entry), 05 備份與還原 (JSON export/import, and XLSX export/import via SheetJS for holdings).
- Charts (`renderYearlyChart`, `renderWeeklyChart`) are rendered with Chart.js against `syncedOverall`/`syncedWeekly` output, so they always reflect live edits, not just the static baseline.

### Cloud (Vercel + Firestore) build

This is the primary phone path and the one the user actually uses day to day. Unlike the Artifact build it needs no human in the loop: `serve.py` pushes a snapshot to Firestore on every write, and the deployed page pulls it after login.

- **Live at `https://handan-web.vercel.app`** (Vercel project `handan-web`), built by `build_web.py` into `handan-web/`, which is gitignored. Deploy with `vercel deploy --prod --yes` from that directory.
- **Firebase project `myfirebase-22090`, shared with the user's separate `Claude_台股分析` project.** Firestore rules live in that repo's `firestore.rules`; the `handan/{docId}` block there grants read only to the owner's uid and denies all client writes. Writes come from `serve.py` via a service account, which bypasses rules entirely. If those rules are ever re-published from that file, the handan block must survive or the phone stops loading.
- **`firebase-service-account.json` sits in `HanDan/` and is gitignored.** Without it `serve.py` prints one line and disables cloud sync; everything else keeps working, so its absence is never fatal.
- **Push is triggered centrally in `Handler._send_json`**: any 200 from a POST, plus the one GET that writes (`updateXlsx=1`). New write endpoints need no extra wiring. Pushes are debounced by `CLOUD_PUSH_DELAY_SEC` since the page mirrors storage on every keystroke-level write.

**The security constraint that shaped this build**: a Vercel URL is world-readable, and Vercel's password protection is a paid feature. The login gate only guards Firestore, so anything left in the HTML is public to anyone who views source. Three blocks therefore must never ship in `handan-web/index.html` — `PORTFOLIO_DATA` (~110 KB, the full ledger since 2021), `STOCK_CODE_MAP` (every holding's name), and `EMBEDDED_PRICES` (every holding's code). `build_web.py` strips all three and aborts if any one of them fails to match, rather than deploying a half-sanitised page. They are re-injected into memory after login from the snapshot's `portfolio` and `codes` keys. **Any new constant carrying holdings data must be added to that strip list.** `STOCK_CODE_MAP` is parsed from the HTML rather than reused from `fetch_close.py` because the two legitimately differ — the HTML has an extra 凱基金 alias for 2883.

**Version check for stale phone caches**: the page is added to the iPhone home screen, so Safari caches it aggressively and a code change can go unnoticed. `build_web.py` therefore writes a `version.json` next to `index.html` and stamps the same id into the page; on every load the page fetches that file and shows an update bar when the two differ. The id is a SHA-256 prefix of the page **computed before the id itself is injected**, so rebuilding unchanged content yields the same id and never produces a phantom update prompt. The generated `vercel.json` marks only `version.json` as `no-store` — the 800 KB main file stays cacheable, which is the whole point of checking a tiny separate file. The update button reloads with `?v=<id>` because a plain reload can be served from cache again.

Two structural details that are easy to break:

- **The app script must not self-start.** `build_web.py` retags it `type="text/handan-app"` so the browser keeps it as inert text; the boot script injects it only after login, after `PORTFOLIO_DATA` and the code map are in place and `localStorage` is preloaded. Starting it earlier means it reads empty state.
- **Closing-tag replacements must run from the end of the file.** The inlined SheetJS contains an HTML export template whose string literals include `</head><body>` and `</body></html>`, and they appear *before* the real tags. `must_replace_last()` exists for exactly this; a forward search silently injects into a JavaScript string and breaks the whole page.

### Working with this codebase

- All UI strings, comments, and docstrings are in Traditional Chinese (Taiwan usage) — match that when editing this project.
- Since the HTML is one file with no source maps or components, use line-anchored greps (e.g. `EMBEDDED_PRICES`, `PORTFOLIO_DATA`, function names above) to navigate rather than assuming a directory structure.
- `PORTFOLIO_DATA` and `EMBEDDED_PRICES` are both large single-line JSON/JS literals — avoid full-file reads; target them with line offsets or regex search instead.
