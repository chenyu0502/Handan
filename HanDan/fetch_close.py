"""
菡萏咖啡．台股收盤價抓取工具

用途：
    在本機抓取上市（TWSE）與上櫃（TPEx）的盤後收盤價，
    輸出可直接貼回「菡萏咖啡-台股損益存摺.html」的 EMBEDDED_PRICES 區塊。

取價日期規則：
    1. 執行時間在 13:30 前（尚未收盤）→ 取前一個交易日。
    2. 執行時間在 13:30 後（已收盤）  → 取當日。
    3. 遇週六、週日自動往前推至最近平日。
    4. 國定假日無法由程式判斷。若當日休市，來源會回傳前一交易日的資料，
       實際日期以輸出的 DATA_DATE 為準。

執行方式（Windows + uv）：
    uv venv
    .venv\\Scripts\\activate
    uv pip install requests
    python fetch_close.py
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import re
import shutil
import ssl
import sys
from datetime import datetime, timedelta
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- 設定 ---------------------------------------------------------------

TWSE_CSV = "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=open_data"
TWSE_DATED = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL?response=json&date={ymd}"
TPEX_JSON = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"
# 上市櫃共用的即時報價引擎。tpex_mainboard_daily_close_quotes 這份官方批次檔
# 收盤後常要等一兩小時才處理完當天資料，但這個引擎在收盤當下（最後成交時間
# 13:30:00）就已經反映最後成交價，可以更快拿到當日收盤價，用來補上還沒更新的部分。
MIS_QUOTE_URL = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
# 即時報價引擎單次查詢的檔數有上限，超過會整批查不到；同一代號要查上市與
# 上櫃兩種前綴，目標數是持股檔數的兩倍，因此分批送出。
MIS_BATCH_SIZE = 40

# 興櫃股票屬櫃買中心另一組資料集，主板端點不涵蓋。
# 端點名稱歷經改版，逐一嘗試直到取得資料。
TPEX_ESB_URLS = [
    "https://www.tpex.org.tw/openapi/v1/tpex_esb_latest_statistics",
    "https://www.tpex.org.tw/openapi/v1/tpex_esb_daily_close_quotes",
    "https://www.tpex.org.tw/openapi/v1/esb_latest_statistics",
]

HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json, text/csv, */*"}
# 證交所的每日完整 CSV 有一千多檔，尖峰時段回應常常拖過 20 秒；逾時會讓
# 整個抓價流程拋例外，網頁端只會看到一個沒有頭緒的「HTTP 500」。放寬秒數
# 並自動重試，把間歇性的慢回應吸收掉。
TIMEOUT = 45
_RETRY = Retry(
    total=2, connect=2, read=2,
    backoff_factor=1.5,
    status_forcelist=(429, 500, 502, 503, 504),
    allowed_methods=frozenset(["GET"]),
)


class _RelaxedCertAdapter(HTTPAdapter):
    """tpex.org.tw 的憑證鏈缺少 Subject Key Identifier 擴充欄位，OpenSSL 3.2+
    預設的嚴格模式（X509_V_FLAG_X509_STRICT）會因此拒絕連線，即便憑證本身
    未過期、主機名稱也相符。這裡僅關閉「嚴格」旗標，其餘驗證（憑證鏈、
    主機名稱）維持正常，範圍只限這個 Session，不影響 TWSE 的請求。"""

    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
        kwargs["ssl_context"] = ctx
        return super().init_poolmanager(*args, **kwargs)


_tpex_session = requests.Session()
_tpex_session.mount("https://www.tpex.org.tw", _RelaxedCertAdapter(max_retries=_RETRY))

# TWSE 與其餘來源共用，統一帶上重試策略
_session = requests.Session()
_session.mount("https://", HTTPAdapter(max_retries=_RETRY))
_session.mount("http://", HTTPAdapter(max_retries=_RETRY))

# 持倉清單。ETF 直接用代號；個股為「中文名: 代號」。
# 代號皆已比對證交所官方檔案的「證券代號／證券名稱」欄位。
ETF_CODES = [
    "0050", "0056", "00403A", "00679B", "00687B", "00713", "00720B",
    "00878", "00888", "00891", "00900", "00904", "00915", "00917",
    "00919", "00929", "00934", "00937B", "00939", "00940", "00942B",
    "00944", "00945B", "00950B", "00980A", "00981A", "00982A",
    "00984A", "00992A",
]

STOCK_CODE_MAP = {
    "大將": "1453", "永光": "1711", "三晃": "1721", "南亞": "1303",
    "和大": "1536", "宏泰": "1612", "億泰": "1616", "華新": "1605",
    "茂矽": "2342", "燿華": "2367", "國建": "2501", "冠德": "2520",
    "聯電": "2303", "廣宇": "2328", "台積電": "2330", "旺宏": "2337",
    "宏碁": "2353", "所羅門": "2359", "矽統": "2363", "大同": "2371",
    "長榮": "2603", "開發金": "2883", "玉山金": "2884", "銘異": "3060",
    "緯創": "3231", "群創": "3481", "炎洲": "4306", "十銓": "4967",
    "達麗": "6177", "力成": "6239", "康舒": "6282", "興能高": "6558",
    "力積電": "6770",
    # 以下 6 檔代號由使用者提供並確認。
    "加百裕": "3323", "家登": "3680", "錸恩帕斯": "7907", "南電": "8046",
    "羅昇": "8374", "國統": "8936",
}

# 尚未確認代號的持股。目前全部已確認，清單為空。
UNRESOLVED: list[str] = []


# --- 日期規則 -----------------------------------------------------------

def target_trading_date(now: datetime | None = None) -> tuple[datetime, bool]:
    """依 13:30 分界與週末規則，回傳應取價的日期與「是否尚未收盤」旗標。"""
    now = now or datetime.now()
    before_close = (now.hour, now.minute) < (13, 30)
    d = now - timedelta(days=1) if before_close else now
    while d.weekday() >= 5:  # 5 = 週六，6 = 週日
        d -= timedelta(days=1)
    return d, before_close


def to_roc(d: datetime) -> str:
    """西元轉民國年格式，例：2026-07-30 -> 1150730。"""
    return f"{d.year - 1911}{d.month:02d}{d.day:02d}"


# --- 抓取 ---------------------------------------------------------------

def fetch_twse_dated(d: datetime) -> tuple[dict[str, float], str]:
    """抓取指定日期的上市收盤價。查無資料時回傳空 dict。

    端點雖然帶 response=json，但證交所已改為一律回傳 CSV（不論參數），
    JSON 解析必定失敗；欄位順序與 fetch_twse_latest 的全量 CSV 相同，
    故沿用同一套解析方式。
    """
    url = TWSE_DATED.format(ymd=d.strftime("%Y%m%d"))
    r = _session.get(url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    r.encoding = "utf-8"
    reader = csv.reader(io.StringIO(r.text))
    next(reader, None)  # 跳過標題列
    prices: dict[str, float] = {}
    data_date = ""
    for row in reader:
        if len(row) < 9:
            continue
        roc_date, code, close = row[0].strip(), row[1].strip(), row[8].strip().replace(",", "")
        try:
            value = float(close)
        except ValueError:
            continue
        if value > 0:
            prices[code] = value
            data_date = data_date or roc_date
    return prices, data_date


def fetch_twse_latest() -> tuple[dict[str, float], str]:
    """抓取上市最新一期收盤價（不指定日期），回傳價格與資料日期。"""
    r = _session.get(TWSE_CSV, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    r.encoding = "utf-8"
    reader = csv.reader(io.StringIO(r.text))
    next(reader, None)  # 跳過標題列
    prices: dict[str, float] = {}
    data_date = ""
    for row in reader:
        if len(row) < 9:
            continue
        roc_date, code, close = row[0].strip(), row[1].strip(), row[8].strip().replace(",", "")
        try:
            value = float(close)
        except ValueError:
            continue
        if value > 0:
            prices[code] = value
            data_date = data_date or roc_date
    return prices, data_date


def roc_to_iso(raw: str) -> str:
    """民國年日期字串（7 或 8 碼）轉西元 ISO 格式；格式不明則原樣回傳。"""
    if raw and len(raw) == 7 and raw.isdigit():
        return f"{int(raw[:3]) + 1911}-{raw[3:5]}-{raw[5:7]}"
    if raw and len(raw) == 8 and raw.isdigit():
        return f"{raw[:4]}-{raw[4:6]}-{raw[6:8]}"
    return raw


def _parse_quotes(rows: list) -> dict[str, float]:
    """從櫃買中心回傳的列表萃取「代號 -> 收盤價」。

    欄位命名因資料集與改版而異，逐一嘗試候選鍵值。
    """
    code_keys = ("SecuritiesCompanyCode", "Code", "code", "股票代號", "代號")
    close_keys = ("Close", "ClosingPrice", "close", "收盤", "收盤價",
                  "LastPrice", "LatestPrice", "最後成交價", "均價", "AvgPrice")
    prices: dict[str, float] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        code = next((str(row[k]).strip() for k in code_keys if row.get(k)), None)
        close = next((str(row[k]).strip().replace(",", "")
                      for k in close_keys if row.get(k)), None)
        if not code or not close:
            continue
        try:
            value = float(close)
        except ValueError:
            continue
        if value > 0:
            prices[code] = value
    return prices


def _extract_date(rows: list) -> str:
    """從櫃買中心回傳列表的第一筆資料取得該批資料的日期（民國年原始格式）。

    櫃買中心的「收盤報價」端點常常沒有即時更新到當天——即使已過收盤
    時間，仍可能回傳前一交易日的資料。回傳這個原始日期供呼叫端比對，
    才能發現「抓到資料但其實是舊的」這種情況，而不是誤以為已是當日收盤價。
    """
    date_keys = ("Date", "日期")
    for row in rows:
        if not isinstance(row, dict):
            continue
        date = next((str(row[k]).strip() for k in date_keys if row.get(k)), None)
        if date:
            return date
    return ""


def fetch_tpex_esb() -> tuple[dict[str, float], str, str]:
    """抓取興櫃收盤價。回傳 (價格, 成功的端點名稱, 該批資料日期)。"""
    for url in TPEX_ESB_URLS:
        try:
            r = _tpex_session.get(url, headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            rows = r.json()
            prices = _parse_quotes(rows)
            if prices:
                return prices, url.rsplit("/", 1)[-1], _extract_date(rows)
        except Exception:  # noqa: BLE001, S112
            continue
    return {}, "", ""


def fetch_tpex() -> tuple[dict[str, float], str]:
    """抓取上櫃收盤價。欄位名稱因來源改版而有多種可能，逐一嘗試。

    回傳 (價格, 該批資料日期)——日期用來偵測來源尚未更新到當天的情況。
    """
    r = _tpex_session.get(TPEX_JSON, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    rows = r.json()
    return _parse_quotes(rows), _extract_date(rows)


def fetch_realtime_quotes(codes: list[str]) -> tuple[dict[str, float], dict[str, float], str]:
    """用上市櫃共用的即時報價引擎取得當日最後成交價與前一交易日收盤價。

    官方每日批次檔要到收盤後一段時間才產出（實測 13:55 仍是前一交易日的
    資料），剛收盤就抓價會拿到昨天的數字；這個引擎在收盤當下就已反映最後
    成交價，用來補上當天的價格。

    同一代號同時以 `tse_<代號>.tw` 與 `otc_<代號>.tw` 兩種前綴查詢，引擎只
    會回傳實際存在的那一筆（另一筆的成交價欄位是 "-"，會被下面的檢查濾掉），
    因此呼叫端不需要先分辨商品屬於上市還是上櫃，把整份追蹤清單丟進來即可。

    單次查詢的檔數有上限，超過會整批查不到，所以分批送出。

    同時取回前一交易日收盤價（欄位 y），這是「今日損益」唯一的來源：官方
    每日檔只有當天的四價，沒有前一日收盤，而 xlsx 的現價欄位也只存一個
    數字，都算不出今天漲跌了多少。

    回傳 (當日價, 前一日收盤價, 資料日期)。
    """
    if not codes:
        return {}, {}, ""

    targets = [f"{market}_{c}.tw" for c in codes for market in ("tse", "otc")]
    prices: dict[str, float] = {}
    prev_closes: dict[str, float] = {}
    date = ""

    for i in range(0, len(targets), MIS_BATCH_SIZE):
        params = {"ex_ch": "|".join(targets[i:i + MIS_BATCH_SIZE]),
                  "json": "1", "delay": "0"}
        r = _session.get(MIS_QUOTE_URL, params=params, headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        # 回應前面夾雜大量空白/換行（防爬蟲雜訊），JSON 本體從第一個 { 開始。
        text = r.text
        start = text.find("{")
        if start < 0:
            continue
        data = json.loads(text[start:])

        for row in data.get("msgArray", []):
            code = row.get("c")
            if not code:
                continue

            # 前一交易日收盤價（y）跟當日有沒有成交無關，先收下來再說。
            # 當日無成交的商品成交價欄位是 "-"，但 y 仍然有值；早期版本在
            # 成交價無效時整筆跳過，連帶把 y 也丟掉，那些商品就完全沒有
            # 比較基準，今日損益會整檔漏算（實測 00944 因此短少 3,700）。
            try:
                prev = float(row.get("y"))
                if prev > 0:
                    prev_closes[code] = prev
            except (TypeError, ValueError):
                pass

            z = row.get("z")  # 最後成交價；收盤後即為當日收盤價。
            if not z or z == "-":
                continue
            try:
                value = float(z)
            except ValueError:
                continue
            if value <= 0:
                continue

            prices[code] = value
            date = date or (row.get("d") or "")

    return prices, prev_closes, date


# --- 輸出 ---------------------------------------------------------------

def build_js_snippet(prices: dict[str, float], iso_date: str) -> str:
    """組出可直接取代 HTML 內 EMBEDDED_PRICES 的 JavaScript 片段。"""
    etf = {c: prices[c] for c in ETF_CODES if c in prices}
    stock = {code: prices[code] for code in sorted(set(STOCK_CODE_MAP.values())) if code in prices}

    def fmt(d: dict[str, float]) -> str:
        items = [f"'{k}': {v:g}" for k, v in d.items()]
        lines, cur = [], "  "
        for it in items:
            if len(cur) + len(it) + 2 > 96:
                lines.append(cur.rstrip())
                cur = "  "
            cur += it + ", "
        if cur.strip():
            lines.append(cur.rstrip().rstrip(","))
        return "\n".join(lines)

    etf_block = fmt(etf)
    stock_block = fmt(stock)
    # 兩個區塊之間必須有逗號，否則貼回 HTML 會造成 JavaScript 語法錯誤。
    if etf_block and not etf_block.rstrip().endswith(","):
        etf_block += ","

    return (
        f"const EMBEDDED_PRICES_DATE = '{iso_date}';\n"
        "const EMBEDDED_PRICES = {\n"
        "  // ETF\n"
        f"{etf_block}\n"
        "  // 個股\n"
        f"{stock_block}\n"
        "};"
    )


def patch_html(html_path: Path, snippet: str) -> tuple[bool, str]:
    """就地改寫 HTML 中的 EMBEDDED_PRICES 區塊。

    回傳 (是否成功, 訊息)。改寫前會先備份為 .bak 檔。
    採用正規表示式定位，不依賴行號，HTML 其餘內容改版後仍可運作。
    """
    if not html_path.exists():
        return False, f"找不到檔案：{html_path}"

    original = html_path.read_text(encoding="utf-8")
    pattern = re.compile(
        r"const EMBEDDED_PRICES_DATE\s*=\s*'[^']*';\s*"
        r"const EMBEDDED_PRICES\s*=\s*\{.*?\n\};",
        re.DOTALL,
    )
    if not pattern.search(original):
        return False, "HTML 中找不到 EMBEDDED_PRICES 區塊，可能檔案版本不符。"

    updated = pattern.sub(lambda _m: snippet, original, count=1)
    if updated == original:
        return False, "取代後內容未變動，已中止以免產生無效寫入。"

    backup = html_path.with_suffix(html_path.suffix + ".bak")
    shutil.copy2(html_path, backup)
    html_path.write_text(updated, encoding="utf-8")
    return True, f"已更新 {html_path.name}，原檔備份為 {backup.name}"


def probe_code(code: str) -> int:
    """逐一查詢三組資料集，回報指定代號出現在哪一組。用於排查缺漏。"""
    code = code.strip()
    print(f"[查詢] 代號 {code} 落在哪一組資料集\n")
    found = False

    try:
        twse, _ = fetch_twse_latest()
        hit = code in twse
        print(f"  上市 TWSE          共 {len(twse):5d} 檔  {'命中 ' + str(twse[code]) if hit else '無'}")
        found |= hit
    except Exception as exc:  # noqa: BLE001
        print(f"  上市 TWSE          查詢失敗：{type(exc).__name__}")

    try:
        tpex, tpex_date = fetch_tpex()
        hit = code in tpex
        print(f"  上櫃 TPEx 主板     共 {len(tpex):5d} 檔（資料日期 {roc_to_iso(tpex_date)}）  "
              f"{'命中 ' + str(tpex[code]) if hit else '無'}")
        found |= hit
    except Exception as exc:  # noqa: BLE001
        print(f"  上櫃 TPEx 主板     查詢失敗：{type(exc).__name__}")

    esb, esb_src, esb_date = fetch_tpex_esb()
    hit = code in esb
    src = f"（{esb_src}，資料日期 {roc_to_iso(esb_date)}）" if esb_src else "（所有端點皆失敗）"
    print(f"  興櫃 TPEx ESB      共 {len(esb):5d} 檔{src}  {'命中 ' + str(esb[code]) if hit else '無'}")
    found |= hit

    print()
    if found:
        print("[結論] 代號有效，程式可正常抓取。")
    else:
        print("[結論] 三組資料集皆查無此代號。請於券商 App 確認代號是否正確，")
        print("       或該商品是否為公開發行前、已下市等不在公開報價範圍的狀態。")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="抓取台股收盤價並可選擇直接更新 HTML")
    parser.add_argument(
        "--html", type=Path, default=None,
        help="要更新的 HTML 檔案路徑。指定後會就地改寫並自動備份。",
    )
    parser.add_argument(
        "--probe", metavar="CODE", default=None,
        help="查詢某代號落在哪一組資料集（上市／上櫃／興櫃），用於排查查無報價的商品。",
    )
    args = parser.parse_args()

    if args.probe:
        return probe_code(args.probe)

    target, before_close = target_trading_date()
    state = "尚未收盤，取前一交易日" if before_close else "已收盤，取當日"
    print(f"[規則] 現在時間 {datetime.now():%Y-%m-%d %H:%M}（{state}）")
    print(f"[規則] 目標取價日：{target:%Y-%m-%d}\n")

    # 先試指定日期，失敗再退回最新一期。
    try:
        twse, raw_date = fetch_twse_dated(target)
        source = f"指定日期 {target:%Y-%m-%d}"
    except Exception as exc:  # noqa: BLE001
        print(f"[警告] 指定日期查詢失敗（{type(exc).__name__}: {exc}），改抓最新一期。")
        twse, raw_date = {}, ""
        source = ""

    if not twse:
        twse, raw_date = fetch_twse_latest()
        source = "最新一期"

    try:
        tpex, tpex_date = fetch_tpex()
    except Exception as exc:  # noqa: BLE001
        print(f"[警告] 上櫃資料抓取失敗（{type(exc).__name__}: {exc}），僅輸出上市資料。")
        tpex, tpex_date = {}, ""

    try:
        esb, esb_src, esb_date = fetch_tpex_esb()
        if esb:
            print(f"[結果] 興櫃 {len(esb)} 檔，端點：{esb_src}")
        else:
            print("[警告] 興櫃資料取得失敗，所有候選端點皆無回應。")
    except Exception as exc:  # noqa: BLE001
        print(f"[警告] 興櫃資料抓取失敗（{type(exc).__name__}: {exc}）")
        esb, esb_date = {}, ""

    prices = {**twse, **tpex, **esb}
    print(f"[結果] 上市 {len(twse)} 檔、上櫃 {len(tpex)} 檔、興櫃 {len(esb)} 檔，來源：{source}")

    iso_date = roc_to_iso(raw_date) or target.strftime("%Y-%m-%d")
    target_iso = target.strftime("%Y-%m-%d")
    tracked = sorted(set(ETF_CODES) | set(STOCK_CODE_MAP.values()))

    print(f"[結果] 實際資料日期：{iso_date}")

    # 官方每日檔要到收盤後一段時間才產出（實測 13:55 仍回前一交易日的資料），
    # 這時上市與上櫃兩邊都是舊的、彼此一致，靠來源互比看不出問題，收盤後隨即
    # 抓價就只會拿到昨天的數字。只要已收盤卻拿到舊日期，就直接用即時報價引擎
    # （收盤當下即反映最後成交價）補上當日價。
    realtime_applied = False
    if not before_close and iso_date != target_iso:
        try:
            rt_prices, _rt_prev, rt_date = fetch_realtime_quotes(tracked)
        except Exception as exc:  # noqa: BLE001
            rt_prices, rt_date = {}, ""
            print(f"[警告] 即時報價備援抓取失敗（{type(exc).__name__}: {exc}）")
        if rt_prices and roc_to_iso(rt_date) == target_iso:
            prices.update(rt_prices)
            print(f"[結果] 官方每日檔仍為 {iso_date}，已用即時報價補上 "
                  f"{len(rt_prices)} 檔 {target_iso} 收盤價。")
            iso_date = target_iso
            realtime_applied = True
        else:
            print(f"[警告] 實際資料日期與目標取價日 {target_iso} 不符，"
                  "可能為國定假日休市，或官方檔與即時報價都尚未更新。")
    elif iso_date != target_iso:
        print(f"[警告] 實際資料日期與目標取價日 {target_iso} 不符，可能為國定假日休市或來源尚未更新。")

    # 上市（TWSE）與上櫃／興櫃（TPEx）是各自獨立的來源，收盤後更新的時間點不一定一致；
    # 只看 TWSE 的日期會誤以為全部資料都是當日的，因此個別比對。
    # 上面若已整批補過當日價就不必再補一次。
    if not realtime_applied and tpex and roc_to_iso(tpex_date) != iso_date:
        try:
            rt_prices, _rt_prev, rt_date = fetch_realtime_quotes(tracked)
        except Exception as exc:  # noqa: BLE001
            rt_prices, rt_date = {}, ""
            print(f"[警告] 即時報價備援抓取失敗（{type(exc).__name__}: {exc}）")
        if rt_prices and roc_to_iso(rt_date) == iso_date:
            prices.update(rt_prices)
            tpex.update(rt_prices)
            print(f"[結果] 上櫃官方批次檔為 {roc_to_iso(tpex_date)}，尚未更新至 {iso_date}，"
                  f"已改用即時報價補上 {len(rt_prices)} 檔。")
        else:
            print(f"[警告] 上櫃資料為 {roc_to_iso(tpex_date)} 收盤價，尚未更新至 {iso_date}，"
                  "上櫃個股／債券型 ETF 價格可能不是最新收盤價（即時報價備援也未取得今日資料）。")
    if esb and roc_to_iso(esb_date) != iso_date:
        print(f"[警告] 興櫃資料為 {roc_to_iso(esb_date)} 收盤價，尚未更新至 {iso_date}。")

    wanted = set(ETF_CODES) | set(STOCK_CODE_MAP.values())
    hit = wanted & prices.keys()
    miss = sorted(wanted - prices.keys())
    print(f"[結果] 持倉命中 {len(hit)}/{len(wanted)} 檔")
    if miss:
        print(f"[結果] 查無報價：{'、'.join(miss)}")
    if UNRESOLVED:
        print(f"[待補] 尚未確認代號，未納入抓取：{'、'.join(UNRESOLVED)}")

    with open("prices.json", "w", encoding="utf-8") as f:
        json.dump(
            {"date": iso_date, "prices": {c: prices[c] for c in sorted(hit)}},
            f, ensure_ascii=False, indent=2,
        )

    snippet = build_js_snippet(prices, iso_date)
    with open("embedded_prices.js", "w", encoding="utf-8") as f:
        f.write(snippet + "\n")

    print("\n[輸出] prices.json（完整價格）")
    print("[輸出] embedded_prices.js（可貼回 HTML 的片段）")

    if args.html:
        ok, msg = patch_html(args.html, snippet)
        print(f"[{'輸出' if ok else '錯誤'}] {msg}")
        if not ok:
            return 1
        print("\n[完成] HTML 已直接更新，開啟即為最新收盤價，無需手動貼上。")
        return 0

    print("\n--- 以下內容取代 HTML 中的 EMBEDDED_PRICES_DATE 與 EMBEDDED_PRICES ---")
    print(snippet)
    return 0


if __name__ == "__main__":
    sys.exit(main())
