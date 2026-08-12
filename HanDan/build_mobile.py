"""把「菡萏咖啡-台股損益存摺.html」轉成可發布為 Claude Artifact 的手機查看版。

手機端不經過 serve.py，也連不到任何外部主機，因此需要一份完全自足的單檔版本。
Artifact 的內容安全性原則會封鎖所有外部請求（CDN 腳本、字型、圖片一律擋掉），
而且發布時頁面內容會被包進既有的 <!doctype html><head></head><body> 骨架，
所以本腳本做四件事：

1. 把 Chart.js 與 SheetJS 由 CDN 連結改為直接內嵌。
2. 移除 Google Fonts 連結，並補強字型 fallback 至 iOS 內建字型。
3. 剝除 DOCTYPE、html、head、body 等結構標籤，只留下內容片段。
4. 以 JS 動態插入 viewport meta，確保手機以裝置寬度渲染。
5. 內嵌 serve.py 三個唯讀 API 的資料快照，並攔截 fetch 供應，
   讓手機顯示的數值與電腦上透過 開啟存摺.bat 看到的一致。

原始檔的 CSS、版面與所有計算邏輯完全不更動，本腳本只處理載入方式。

用法（一般情況請直接雙擊「更新手機版.bat」，不必手動執行）：

    python build_mobile.py
"""

import json
import re
from pathlib import Path

import fetch_close as fc

BASE = Path(__file__).resolve().parent
SRC = BASE / "菡萏咖啡-台股損益存摺.html"
OUT = BASE / "菡萏咖啡-手機版.html"
VENDOR = BASE / "vendor"

CHART_CDN = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>'
XLSX_CDN = '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'

FONT_PRECONNECT = '<link rel="preconnect" href="https://fonts.googleapis.com">'
FONT_LINK = (
    '<link href="https://fonts.googleapis.com/css2?'
    "family=Noto+Serif+TC:wght@500;600;700;900&"
    "family=Noto+Sans+TC:wght@400;500;700&"
    'family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">'
)

# 沒有 Google Fonts 可用時，改用 iOS 與 Windows 內建字型遞補
FONT_FALLBACKS = [
    ("'Noto Sans TC',sans-serif", "'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif"),
    ("'Noto Serif TC',serif", "'Noto Serif TC','Songti TC','PingFang TC',serif"),
    ("'IBM Plex Mono',monospace", "'IBM Plex Mono','SF Mono',Menlo,Consolas,monospace"),
]

# Artifact 骨架未必帶 viewport meta，手機若缺此設定會以桌面寬度渲染
VIEWPORT_SHIM = """<script>
(function () {
  if (document.querySelector('meta[name="viewport"]')) return;
  var m = document.createElement('meta');
  m.name = 'viewport';
  m.content = 'width=device-width, initial-scale=1.0';
  document.head.appendChild(m);
})();
</script>
"""

# 手機版沒有 serve.py 可連，頁面對 /api/ 的請求會全部落空，畫面就只能退回
# PORTFOLIO_DATA 的舊靜態基準，數值與電腦上看到的對不起來。這裡在建置當下
# 先把 xlsx 讀成快照內嵌，再攔截 fetch 直接回覆，原始 HTML 一行都不用改。
API_SHIM = """<script>
(function () {
  var SNAPSHOT = __SNAPSHOT__;
  var LS_PREFIX = '__LS_PREFIX__';
  var realFetch = (typeof window.fetch === 'function') ? window.fetch.bind(window) : null;

  // 告訴頁面自己是手機版，storage 寫入時就不會再嘗試回傳鏡像給 serve.py
  window.__HANDAN_MOBILE__ = true;

  // 把電腦上的手動編輯（持股調整、記一筆、確認並存檔等）預載進本機儲存。
  // 以 savedAt 當版本記號：同一份快照只覆蓋一次，之後你在手機上改的東西
  // 不會每次重新整理就被蓋掉；等電腦端重新發布新快照時才會再次覆蓋。
  (function preloadState() {
    var st = SNAPSHOT.state;
    if (!st || !st.ok || !st.data) return;
    try {
      var MARK = LS_PREFIX + '__mirror_stamp__';
      if (localStorage.getItem(MARK) === st.savedAt) return;
      for (var k in st.data) {
        if (Object.prototype.hasOwnProperty.call(st.data, k)) {
          localStorage.setItem(LS_PREFIX + k, st.data[k]);
        }
      }
      localStorage.setItem(MARK, st.savedAt);
    } catch (e) {
      // 無痕模式等情境下 localStorage 不可寫，略過即可
    }
  })();

  function reply(data) {
    var body = JSON.stringify(data);
    if (typeof Response === 'function') {
      return Promise.resolve(new Response(body, {
        status: 200,
        headers: { 'Content-Type': 'application/json' }
      }));
    }
    // 極舊版瀏覽器沒有 Response 建構子時的最小替代品
    return Promise.resolve({
      ok: true,
      status: 200,
      json: function () { return Promise.resolve(JSON.parse(body)); },
      text: function () { return Promise.resolve(body); }
    });
  }

  window.fetch = function (input, init) {
    var url = '';
    try {
      url = (typeof input === 'string') ? input : ((input && input.url) || '');
    } catch (e) {
      url = '';
    }

    if (url.indexOf('/api/xlsx-prices') !== -1) return reply(SNAPSHOT.prices);
    if (url.indexOf('/api/xlsx-balance') !== -1) return reply(SNAPSHOT.balance);
    if (url.indexOf('/api/xlsx-weekly') !== -1) return reply(SNAPSHOT.weekly);
    if (url.indexOf('/api/') !== -1) {
      // 其餘都是寫回 Excel 的動作，手機版只能查看
      return reply({ ok: false, error: '手機版為唯讀，請在電腦上操作後重新產生手機版。' });
    }

    return realFetch ? realFetch(input, init)
                     : Promise.reject(new Error('此環境不支援 fetch'));
  };
})();
</script>
"""

# 與 fetch_close.patch_html 相同的定位規則。這裡是對字串操作而非改寫檔案
# （不需要備份、也不寫回主檔），因此複用規則而不呼叫那支函式。
EMBEDDED_PRICES_PATTERN = re.compile(
    r"const EMBEDDED_PRICES_DATE\s*=\s*'[^']*';\s*"
    r"const EMBEDDED_PRICES\s*=\s*\{.*?\n\};",
    re.DOTALL,
)

# 這些標籤由 Artifact 骨架提供，片段裡不能重複出現
STRIP_TAGS = (
    "<!DOCTYPE html>\n",
    '<html lang="zh-Hant">\n',
    "<head>\n",
    '<meta charset="UTF-8">\n',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n',
    "</head>\n",
    "<body>\n",
    "</body>\n",
    "</html>\n",
)


def must_replace(text: str, old: str, new: str, label: str) -> str:
    """執行替換並確認確實命中，避免來源檔改版後靜默失敗。"""
    if old not in text:
        raise SystemExit(f"[錯誤] 找不到「{label}」的目標字串，來源檔格式可能已變更，轉換中止。")
    return text.replace(old, new, 1)


def read_vendor(name: str) -> str:
    """讀取內嵌用的函式庫，順便擋掉會提前結束 script 區塊的字串。"""
    path = VENDOR / name
    if not path.exists():
        raise SystemExit(f"[錯誤] 缺少 {path}，無法內嵌。請確認 vendor 目錄完整。")
    return path.read_text(encoding="utf-8").replace("</script>", "<\\/script>")


def build_api_snapshot() -> tuple[dict, str, str]:
    """把 serve.py 三個唯讀 API 的結果讀成快照，供手機版離線使用。

    直接重用 serve.py 的函式而不是打 HTTP，這樣不必先啟動伺服器。
    任何一項讀取失敗都不中斷建置：該項標記為 ok=False，頁面會自行退回
    PORTFOLIO_DATA 的靜態基準，行為與連不到 serve.py 時完全一致。

    回傳 (快照 dict, xlsx 狀態描述, 編輯鏡像狀態描述)。回傳 dict 而非
    字面量，是因為呼叫端還要用其中的現價去同步內建備援。
    """
    readers = (
        ("prices", "現價"),
        ("balance", "帳戶餘額"),
        ("weekly", "每週記帳"),
    )
    snapshot: dict[str, dict] = {}
    failed: list[str] = []

    try:
        import serve
    except Exception as exc:  # noqa: BLE001
        for key, _ in readers:
            snapshot[key] = {"ok": False, "error": f"無法載入 serve.py：{exc}"}
        snapshot["state"], state_note = _read_state_mirror()
        return snapshot, f"未內嵌（無法載入 serve.py：{exc}）", state_note

    fns = {
        "prices": serve.read_xlsx_prices,
        "balance": serve.read_xlsx_balance,
        "weekly": serve.read_xlsx_weekly,
    }
    for key, label in readers:
        try:
            snapshot[key] = fns[key]()
        except Exception as exc:  # noqa: BLE001
            snapshot[key] = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
            failed.append(label)

    snapshot["state"], state_note = _read_state_mirror()

    if failed:
        return snapshot, f"部分失敗（{'、'.join(failed)}），該部分將退回靜態基準", state_note

    year = snapshot["weekly"].get("year", "?")
    return snapshot, f"已內嵌（每週記帳年度 {year}）", state_note


def sync_embedded_prices(html: str, prices_snapshot: dict) -> tuple[str, str]:
    """用 Excel 快照的現價改寫手機版內建的備援價格 EMBEDDED_PRICES。

    這份備援只有在 Excel 快照讀取失敗時才會被用到，但若它停在舊日期，那條
    退路就會顯示過期價格。既然快照已經在手上，直接沿用即可，不必再連一次
    證交所。只改寫輸出檔，來源 HTML 不動，維持「唯讀來源、單一輸出」。
    """
    if not prices_snapshot.get("ok"):
        return html, "略過（Excel 快照不可用，保留原有內建價格）"

    prices = prices_snapshot.get("prices") or {}
    if not prices:
        return html, "略過（快照沒有價格資料）"

    # 快照的 date 形如 "2026-08-12 14:13"，內建常數只取日期部分
    date = (prices_snapshot.get("date") or "").split(" ")[0]
    if not date:
        return html, "略過（快照沒有日期）"

    if not EMBEDDED_PRICES_PATTERN.search(html):
        raise SystemExit("[錯誤] 找不到 EMBEDDED_PRICES 區塊，來源檔格式可能已變更，轉換中止。")

    snippet = fc.build_js_snippet(prices, date)
    return EMBEDDED_PRICES_PATTERN.sub(lambda _m: snippet, html, count=1), f"已同步為 {date}（{len(prices)} 檔）"


def _read_state_mirror() -> tuple[dict, str]:
    """讀取 serve.py 寫下的瀏覽器編輯鏡像 mobile_state.json。

    這份檔案由網頁在每次寫入 storage 時自動回傳給 serve.py，是手機版唯一
    拿得到「使用者在電腦上手動修改」的來源。沒有這份檔案不算錯誤，只是
    手機上會少掉那些編輯，因此僅回報狀態、不中斷建置。
    """
    path = BASE / "mobile_state.json"
    if not path.exists():
        return (
            {"ok": False, "error": "尚未有鏡像資料"},
            "未內嵌（請先用 開啟存摺.bat 開一次網頁，讓它把編輯內容同步過來）",
        )
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}, f"讀取失敗：{exc}"

    data["ok"] = True
    return data, f"已內嵌（{data.get('count', 0)} 筆，同步於 {data.get('savedAt', '?')}）"


def _detect_ls_prefix(html: str) -> str:
    """從來源檔取出 localStorage 的鍵前綴，避免日後前綴改動時兩邊對不上。"""
    marker = "const LS_PREFIX = '"
    if marker not in html:
        raise SystemExit("[錯誤] 找不到 LS_PREFIX 定義，無法預載編輯內容，轉換中止。")
    return html.split(marker, 1)[1].split("'", 1)[0]


def _to_js_literal(data: dict) -> str:
    """轉成可安全嵌進 <script> 的 JS 物件字面量。"""
    # 跳脫 </ 以免資料裡剛好出現 </script> 而提前結束區塊
    return json.dumps(data, ensure_ascii=False).replace("</", "<\\/")


def main() -> int:
    if not SRC.exists():
        raise SystemExit(f"[錯誤] 找不到來源檔 {SRC}")

    html = SRC.read_text(encoding="utf-8")

    # 1. 內嵌兩支函式庫。
    #    SheetJS 刻意使用 core 版而非 full 版：full 版內建的 codepage 對應表
    #    含五萬多個 U+FFFD 填充字元，會被 Artifact 部署 API 判定為無效字元而拒絕。
    #    core 版不含該表，但 App 用到的匯出匯入 API 一應俱全，體積還少 444 KB。
    html = must_replace(html, CHART_CDN, f"<script>\n{read_vendor('chart.umd.min.js')}\n</script>", "Chart.js CDN")
    html = must_replace(html, XLSX_CDN, f"<script>\n{read_vendor('xlsx.core.min.js')}\n</script>", "SheetJS CDN")

    # 2. 移除 Google Fonts 並補強字型遞補順序
    html = must_replace(html, FONT_PRECONNECT + "\n", "", "Google Fonts preconnect")
    html = must_replace(html, FONT_LINK + "\n", "", "Google Fonts 樣式表")
    for old, new in FONT_FALLBACKS:
        html = html.replace(old, new)

    # 3. 剝除結構標籤（保留 title，Artifact 會拿它當頁面名稱）
    for tag in STRIP_TAGS:
        html = must_replace(html, tag, "", tag.strip())

    # 4. 讀取 Excel 快照與編輯鏡像
    snapshot, snapshot_note, state_note = build_api_snapshot()

    # 5. 用快照的現價一併更新內建備援，免得快照讀不到時退回過期價格
    html, price_note = sync_embedded_prices(html, snapshot.get("prices", {}))

    # 6. 於開頭補上 viewport 修補、快照與編輯鏡像（必須早於 App 的 script）
    shim = API_SHIM.replace("__SNAPSHOT__", _to_js_literal(snapshot)) \
                   .replace("__LS_PREFIX__", _detect_ls_prefix(html))
    html = VIEWPORT_SHIM + shim + html

    # 發布前把關：U+FFFD 會讓 Artifact 部署 API 直接回 400
    bad = html.count("�")
    if bad:
        raise SystemExit(f"[錯誤] 輸出含 {bad:,} 個無效字元（U+FFFD），發布必定失敗，轉換中止。")

    OUT.write_text(html, encoding="utf-8")

    # 回報目前內嵌的價格日期，方便確認抓價是否生效
    date = "未知"
    marker = "const EMBEDDED_PRICES_DATE = '"
    if marker in html:
        date = html.split(marker, 1)[1].split("'", 1)[0]

    print(f"[完成] 已產生 {OUT.name}")
    print(f"       檔案大小：{len(html.encode('utf-8')):,} bytes")
    print(f"       內嵌價格日期：{date}")
    print(f"       內建備援價格：{price_note}")
    print(f"       Excel 資料快照：{snapshot_note}")
    print(f"       網頁編輯內容　：{state_note}")
    print()
    print("接下來請告訴 Claude「更新手機版」，由 Claude 重新發布到原本的網址。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
