"""把「菡萏咖啡-台股損益存摺.html」轉成部署到 Vercel 的雲端查看版。

與 build_mobile.py（Artifact 版）的差別在資料來源：Artifact 版把快照直接
內嵌在檔案裡，每次更新都要重新發布；這一版改從 Firestore 讀取，PC 端的
serve.py 一有異動就自動推上去，手機打開就是最新的，不需要任何人介入。

因為要連 Firestore，頁面必須先登入再取資料，而登入是非同步的、App 卻是
同步啟動。若照原樣載入，App 會在資料到位前就先渲染，讀到的編輯內容會是
空的。解法是把 App 的 script 標成瀏覽器不認得的 type 讓它不要自動執行，
等登入完成、資料就緒、編輯內容也寫回瀏覽器儲存之後，再動態插入啟動。
如此 App 本身一行都不用改。

用法：

    python build_web.py

輸出 web/index.html，接著用 Vercel 部署 web/ 目錄即可。
"""

import re
from pathlib import Path

BASE = Path(__file__).resolve().parent
SRC = BASE / "菡萏咖啡-台股損益存摺.html"
# 目錄名會被 Vercel 當成專案名稱，也會出現在預設網址裡，因此取一個明確的
OUT_DIR = BASE / "handan-web"
OUT = OUT_DIR / "index.html"
VENDOR = BASE / "vendor"

CHART_CDN = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>'
XLSX_CDN = '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'

# App 程式碼的起點。整份 HTML 只有兩個 inline script，這是後面那個。
APP_SCRIPT_OPEN = "<script>\n/* ============ 菡萏咖啡．台股損益存摺 ============ */"

# 前面那個 script 是 PORTFOLIO_DATA，內含 2021 年至今的完整損益、持股與
# 借券紀錄，約 110 KB。Vercel 部署的網址是公開的，這段若留在檔案裡，任何
# 人不必登入、直接看網頁原始碼就能取得整本帳。因此整段移除，改由登入後
# 從 Firestore 取得並注入。移除後檔案裡不再有任何一個數字。
PORTFOLIO_BLOCK = re.compile(
    r"<script>\s*const PORTFOLIO_DATA = \{.*?\};\s*</script>\s*",
    re.S,
)

# 個股名稱對照表同樣會把投資組合攤開，改成從雲端注入。
CODE_MAP_BLOCK = re.compile(r"const STOCK_CODE_MAP = \{.*?\n\};", re.S)

# 內建備援價格含全部持股代號。雲端版取不到 Firestore 就不會啟動 App，
# 這份備援永遠用不到，直接清空。
EMBEDDED_PRICES_BLOCK = re.compile(r"const EMBEDDED_PRICES = \{.*?\n\};", re.S)

# Firebase 的 web 設定值會打包進前端，這本來就是公開資訊；真正的存取控制
# 在 Firestore 安全規則（只允許指定 uid 讀取、且一律禁止前端寫入）。
FIREBASE_CONFIG = """{
    apiKey: "AIzaSyA2EVtB3J4VCS-NiEcCx2VAVhKFZo38ZO0",
    authDomain: "myfirebase-22090.firebaseapp.com",
    projectId: "myfirebase-22090",
    storageBucket: "myfirebase-22090.firebasestorage.app",
    messagingSenderId: "1023759611635",
    appId: "1:1023759611635:web:8f1710d3e3651adaeae986"
  }"""

# 前端這道檢查只是為了給出友善訊息；真正擋下他人存取的是 Firestore 規則。
ALLOWED_UID = "3U816AjRimMRL7l4x79Z6l2mFtx2"

FIREBASE_SDK = """<script src="https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.14.1/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore-compat.js"></script>
"""

GATE_MARKUP = """<div id="handan-gate">
  <div class="gate-card">
    <div class="gate-title">菡萏咖啡．台股損益存摺</div>
    <div class="gate-sub" id="gateMsg">請先登入以讀取你的資料</div>
    <button id="gateBtn" class="gate-btn">使用 Google 登入</button>
    <div class="gate-note" id="gateNote"></div>
  </div>
</div>
<style>
#handan-gate{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;
  background:#0f1c30;padding:24px;}
#handan-gate.done{display:none;}
.gate-card{background:#f7f2e4;border-radius:14px;padding:32px 28px;max-width:360px;width:100%;
  text-align:center;box-shadow:0 18px 50px rgba(0,0,0,.35);}
.gate-title{font-family:'Noto Serif TC','Songti TC',serif;font-size:19px;font-weight:700;color:#22314f;}
.gate-sub{margin-top:10px;font-size:13px;color:#5b6785;line-height:1.7;}
.gate-btn{margin-top:20px;width:100%;padding:11px 16px;border:none;border-radius:999px;cursor:pointer;
  background:#152647;color:#f7f2e4;font-size:14px;font-weight:600;letter-spacing:.04em;}
.gate-btn:hover{background:#1e3358;}
.gate-btn:disabled{opacity:.55;cursor:default;}
.gate-note{margin-top:14px;font-size:11.5px;color:#8a93a8;line-height:1.6;min-height:1em;}
</style>
"""

BOOT_SCRIPT = """<script>
(function () {
  var CONFIG = __CONFIG__;
  var ALLOWED_UID = '__UID__';
  var DOC_PATH = 'handan/snapshot';

  var gate    = document.getElementById('handan-gate');
  var gateBtn = document.getElementById('gateBtn');
  var gateMsg = document.getElementById('gateMsg');
  var gateNote= document.getElementById('gateNote');

  function note(text) { gateNote.textContent = text || ''; }
  function msg(text)  { gateMsg.textContent = text; }

  // App 啟動後才會呼叫這些 API，屆時快照已經在手上，因此這裡用一個
  // Promise 當閘門：先安裝攔截器，等資料就緒再放行。
  var resolveSnapshot;
  var snapshotReady = new Promise(function (r) { resolveSnapshot = r; });

  function reply(data) {
    var body = JSON.stringify(data);
    if (typeof Response === 'function') {
      return new Response(body, { status: 200, headers: { 'Content-Type': 'application/json' } });
    }
    return { ok: true, status: 200,
             json: function () { return Promise.resolve(JSON.parse(body)); },
             text: function () { return Promise.resolve(body); } };
  }

  var realFetch = (typeof window.fetch === 'function') ? window.fetch.bind(window) : null;
  window.__HANDAN_MOBILE__ = true;   // 唯讀版，不要把狀態鏡像回本機服務

  window.fetch = function (input, init) {
    var url = '';
    try {
      url = (typeof input === 'string') ? input : ((input && input.url) || '');
    } catch (e) { url = ''; }

    if (url.indexOf('/api/xlsx-prices') !== -1)  return snapshotReady.then(function (s) { return reply(s.prices);  });
    if (url.indexOf('/api/xlsx-balance') !== -1) return snapshotReady.then(function (s) { return reply(s.balance); });
    if (url.indexOf('/api/xlsx-weekly') !== -1)  return snapshotReady.then(function (s) { return reply(s.weekly);  });
    if (url.indexOf('/api/') !== -1) {
      return Promise.resolve(reply({ ok: false, error: '雲端版為唯讀，請在電腦上操作。' }));
    }
    return realFetch ? realFetch(input, init)
                     : Promise.reject(new Error('此環境不支援 fetch'));
  };

  // 把電腦上的手動編輯寫回瀏覽器儲存。必須趕在 App 啟動之前完成，
  // 因為 App 一開始就會讀這些鍵，晚一步就會讀到空的。
  function preloadState(state) {
    if (!state || !state.ok || !state.data) return;
    try {
      for (var k in state.data) {
        if (Object.prototype.hasOwnProperty.call(state.data, k)) {
          localStorage.setItem('handan:' + k, state.data[k]);
        }
      }
    } catch (e) { /* 無痕模式等情境下不可寫，略過 */ }
  }

  function startApp() {
    var holder = document.getElementById('handan-app');
    if (!holder) return;
    var s = document.createElement('script');
    s.textContent = holder.textContent;
    document.body.appendChild(s);
  }

  firebase.initializeApp(CONFIG);
  var auth = firebase.auth();
  var db   = firebase.firestore();

  gateBtn.addEventListener('click', function () {
    gateBtn.disabled = true;
    note('正在開啟 Google 登入…');
    auth.signInWithPopup(new firebase.auth.GoogleAuthProvider()).catch(function (err) {
      gateBtn.disabled = false;
      note('登入失敗：' + (err && err.message ? err.message : err));
    });
  });

  auth.onAuthStateChanged(function (user) {
    if (!user) { msg('請先登入以讀取你的資料'); gateBtn.disabled = false; return; }

    if (user.uid !== ALLOWED_UID) {
      msg('這個帳號沒有存取權');
      note('目前登入的是 ' + (user.email || user.uid) + '，請改用存摺擁有者的帳號。');
      gateBtn.textContent = '換一個帳號登入';
      gateBtn.disabled = false;
      auth.signOut();
      return;
    }

    msg('登入成功，正在讀取資料…');
    gateBtn.disabled = true;
    note('');

    db.doc(DOC_PATH).get().then(function (doc) {
      if (!doc.exists) throw new Error('雲端尚無資料，請先在電腦上開啟存摺同步一次');
      var payload = doc.data().payload;
      var snapshot = JSON.parse(payload);

      // 歷史損益基準沒有內嵌在檔案裡（避免未登入就被看光），一定要從雲端
      // 拿到才有辦法啟動，缺了就直接報錯，不要讓 App 帶著空資料跑起來。
      if (!snapshot.portfolio || !snapshot.portfolio.ok || !snapshot.portfolio.data) {
        throw new Error('雲端資料缺少歷史損益基準，請在電腦上開啟存摺重新同步一次');
      }
      if (!snapshot.codes || !snapshot.codes.ok || !snapshot.codes.data) {
        throw new Error('雲端資料缺少股票代號對照表，請在電腦上開啟存摺重新同步一次');
      }
      window.PORTFOLIO_DATA = snapshot.portfolio.data;
      window.__HANDAN_CODES__ = snapshot.codes.data;

      preloadState(snapshot.state);
      resolveSnapshot(snapshot);
      startApp();
      gate.classList.add('done');
    }).catch(function (err) {
      msg('讀取資料失敗');
      note((err && err.message ? err.message : err) + '');
      gateBtn.textContent = '重試';
      gateBtn.disabled = false;
    });
  });
})();
</script>
"""


def must_replace(text: str, old: str, new: str, label: str) -> str:
    """執行替換並確認確實命中，避免來源檔改版後靜默失敗。"""
    if old not in text:
        raise SystemExit(f"[錯誤] 找不到「{label}」的目標字串，來源檔格式可能已變更，轉換中止。")
    return text.replace(old, new, 1)


def must_replace_last(text: str, old: str, new: str, label: str) -> str:
    """從檔尾往回找並替換最後一次出現的位置。

    內嵌的 SheetJS 帶有 HTML 匯出模板，其字串常數裡就含有 </body></html>，
    位置還比真正的結尾標籤更前面。若從頭替換會插進 JavaScript 字串中間，
    把整份檔案弄壞，因此結尾標籤一律從後面找。
    """
    idx = text.rfind(old)
    if idx < 0:
        raise SystemExit(f"[錯誤] 找不到「{label}」的目標字串，來源檔格式可能已變更，轉換中止。")
    return text[:idx] + new + text[idx + len(old):]


def read_vendor(name: str) -> str:
    """讀取內嵌用的函式庫，順便擋掉會提前結束 script 區塊的字串。"""
    path = VENDOR / name
    if not path.exists():
        raise SystemExit(f"[錯誤] 缺少 {path}，無法內嵌。請確認 vendor 目錄完整。")
    return path.read_text(encoding="utf-8").replace("</script>", "<\\/script>")


def main() -> int:
    if not SRC.exists():
        raise SystemExit(f"[錯誤] 找不到來源檔 {SRC}")

    html = SRC.read_text(encoding="utf-8")

    # 1. 內嵌函式庫。Vercel 沒有 Artifact 那種外部主機限制，用 CDN 也行，
    #    但內嵌可以少兩個外部相依，載入也快一點。
    html = must_replace(html, CHART_CDN, f"<script>\n{read_vendor('chart.umd.min.js')}\n</script>", "Chart.js CDN")
    html = must_replace(html, XLSX_CDN, f"<script>\n{read_vendor('xlsx.core.min.js')}\n</script>", "SheetJS CDN")

    # 2. 拿掉所有會洩漏投資組合的資料。這三段都是未登入就看得到的明文，
    #    缺一不可，因此任何一段沒命中都直接中止，不容許只做一半就部署。
    html, n = PORTFOLIO_BLOCK.subn("", html, count=1)
    if n != 1:
        raise SystemExit("[錯誤] 找不到 PORTFOLIO_DATA 區塊，無法確保資料不外洩，轉換中止。")

    html, n = CODE_MAP_BLOCK.subn(
        "const STOCK_CODE_MAP = window.__HANDAN_CODES__ || {};", html, count=1)
    if n != 1:
        raise SystemExit("[錯誤] 找不到 STOCK_CODE_MAP 區塊，無法確保資料不外洩，轉換中止。")

    html, n = EMBEDDED_PRICES_BLOCK.subn("const EMBEDDED_PRICES = {};", html, count=1)
    if n != 1:
        raise SystemExit("[錯誤] 找不到 EMBEDDED_PRICES 區塊，無法確保資料不外洩，轉換中止。")

    # 3. 讓 App 的 script 不要自動執行。type 用瀏覽器不認得的值，內容會被
    #    當成純文字保留，等登入完成後再取出來動態插入。
    html = must_replace(
        html,
        APP_SCRIPT_OPEN,
        '<script id="handan-app" type="text/handan-app">\n'
        "/* ============ 菡萏咖啡．台股損益存摺 ============ */",
        "App script 起點",
    )

    # 4. 在 body 開頭放登入畫面，並在 App script 之前插入啟動程式
    boot = (FIREBASE_SDK
            + BOOT_SCRIPT.replace("__CONFIG__", FIREBASE_CONFIG).replace("__UID__", ALLOWED_UID))
    html = must_replace_last(html, "<body>\n", "<body>\n" + GATE_MARKUP, "body 起始標籤")
    html = must_replace_last(html, "</body>", boot + "</body>", "body 結束標籤")

    OUT_DIR.mkdir(exist_ok=True)
    OUT.write_text(html, encoding="utf-8")

    print(f"[完成] 已產生 {OUT.relative_to(BASE)}")
    print(f"       檔案大小：{len(html.encode('utf-8')):,} bytes")
    print(f"       資料來源：Firestore {FIREBASE_CONFIG.splitlines()[3].strip()}")
    print()
    print("接著部署 web/ 目錄到 Vercel 即可。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
