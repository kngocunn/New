"""
enrich.py — HTTP fast mode (curl_cffi) + Chrome warm-up, SOCKS5 proxy xoay IP.
Enrich SĐT / Email / Năm thành lập / Ngành nghề từ doanh_nghiep.csv.

Cài:  pip install DrissionPage curl_cffi unidecode pandas openpyxl beautifulsoup4
Chạy: python enrich.py
"""

import re, os, sys, time, random, logging, signal, threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock, Event
from unidecode import unidecode
import pandas as pd
from curl_cffi import requests
from bs4 import BeautifulSoup
from DrissionPage import ChromiumPage, ChromiumOptions

import io
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("enrich_hsctvn.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# ─── Config ──────────────────────────────────────────────────────────────────

BASE_URL = "https://hsctvn.com"

DATA_DIR     = r"D:\MCNA\data"
os.makedirs(DATA_DIR, exist_ok=True)
INPUT_PATH   = os.path.join(DATA_DIR, "doanh_nghiep.csv")
CSV_PATH     = os.path.join(DATA_DIR, "doanh_nghiep.csv")
EXCEL_PATH   = os.path.join(DATA_DIR, "doanh_nghiep.xlsx")
CKPT_PATH    = os.path.join(DATA_DIR, "enrich_checkpoint.csv")
LISTING_PATH = os.path.join(DATA_DIR, "listing.csv")

# ── Tốc độ ────────────────────────────────────────────────────────────────────
PARALLEL       = 4     # số thread HTTP đồng thời
DELAY_MIN      = 2   # giây nghỉ ngắn giữa mỗi request
DELAY_MAX      = 2.5
BURST_EVERY    = 40    # nghỉ dài sau mỗi N request
BURST_REST_MIN = 5.0   # giây nghỉ dài
BURST_REST_MAX = 10.0
COOKIE_REFRESH = 60   # refresh CF cookie sau mỗi N request

# ── IP rotation ───────────────────────────────────────────────────────────────
USE_PROXY       = True
LOCAL_PROXY     = "socks5h://127.0.0.1:20000"  # curl_cffi dùng (h = proxy resolve DNS)
CHROME_PROXY    = "socks5://127.0.0.1:20000"   # Chrome dùng (không hỗ trợ socks5h)
IP_ROTATE_EVERY = 60   # giây proxy.exe xoay IP — watcher reload Chrome sau đây

# ── Khác ──────────────────────────────────────────────────────────────────────
RETRY_INCOMPLETE = True  # enrich lại record thiếu cả SDT lẫn Email
FAIL_BEFORE_RECOVER = 5  # số lần HTTP thất bại liên tiếp → trigger recover

# Bắt đầu từ record thứ N trong danh sách chờ (1 = theo checkpoint)
ENRICH_FROM = 1

# ─── Stop ────────────────────────────────────────────────────────────────────

_STOP = Event()

def _handle_sigint(sig, frame):
    if not _STOP.is_set():
        print("\n" + "=" * 60)
        print("  [!] Nhan Ctrl+C — se dung sau record hien tai...")
        print("  [!] Nhan Ctrl+C lan 2 de THOAT NGAY.")
        print("=" * 60 + "\n")
        _STOP.set()
    else:
        print("\n[!] Buoc dung ngay!")
        sys.exit(1)

signal.signal(signal.SIGINT, _handle_sigint)

# ─── Browser (Chrome warm-up + cookie source) ────────────────────────────────

_page: ChromiumPage = None

def get_browser() -> ChromiumPage:
    global _page
    if _page is not None:
        return _page
    log.info("Khoi dong Chrome ...")
    co = ChromiumOptions()
    co.set_argument("--no-sandbox")
    co.set_argument("--disable-blink-features=AutomationControlled")
    co.set_argument("--lang=vi-VN")
    if USE_PROXY:
        co.set_argument(f"--proxy-server={CHROME_PROXY}")
        log.info(f"Chrome dung proxy: {CHROME_PROXY}")
    _page = ChromiumPage(co)
    return _page

def close_browser():
    global _page
    try:
        if _page:
            _page.quit()
    except Exception:
        pass
    _page = None

def _chrome_load_home():
    """Điều hướng Chrome về HOME, chờ trang load xong (vượt CF)."""
    br = get_browser()
    try:
        br.get(BASE_URL)
        t0 = time.time()
        while time.time() - t0 < 60:
            if br.ele("css:ul.hsct", timeout=0) or br.ele("css:ul.hsdn", timeout=0):
                break
            try:
                title = br.title or ""
                if not any(s in title for s in ("Just a moment", "Verify", "Checking")):
                    break
            except Exception:
                pass
            time.sleep(1.5)
    except Exception as e:
        log.warning(f"Chrome load HOME loi: {e}")

# ─── HTTP session ─────────────────────────────────────────────────────────────

_http_session: requests.Session | None = None
_session_lock      = Lock()
_recover_lock      = Lock()
_last_recover_time = 0.0
_request_count     = 0
_count_lock        = Lock()

def init_http_session() -> bool:
    """Copy cookies từ Chrome sang curl_cffi session."""
    global _http_session
    br = get_browser()
    s  = requests.Session(impersonate="chrome120")
    proxies = {"http": LOCAL_PROXY, "https": LOCAL_PROXY} if USE_PROXY else None
    if proxies:
        s.proxies = proxies

    n = 0
    try:
        for c in br.cookies():
            dom = c.get("domain", "")
            if "hsctvn" in dom:
                s.cookies.set(c["name"], c["value"], domain=dom, path=c.get("path", "/"))
                n += 1
    except Exception as e:
        log.warning(f"Khong lay duoc cookies: {e}")
        return False

    if n == 0:
        log.warning("Chua co cookie hsctvn — Chrome chua pass CF?")
        return False

    try:
        ua = br.run_js("return navigator.userAgent;")
    except Exception:
        ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

    s.headers.update({
        "User-Agent":                ua,
        "Accept":                    "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language":           "vi-VN,vi;q=0.9,en;q=0.8",
        "Referer":                   BASE_URL + "/",
        "Upgrade-Insecure-Requests": "1",
        "sec-ch-ua":                 '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
        "sec-ch-ua-mobile":          "?0",
        "sec-ch-ua-platform":        '"Windows"',
        "sec-fetch-site":            "same-origin",
        "sec-fetch-mode":            "navigate",
        "sec-fetch-dest":            "document",
    })
    _http_session = s
    log.info(f"HTTP session OK: {n} cookies")
    return True

def _do_recover(reason: str, url: str = "") -> bool:
    """Reload Chrome + reinit HTTP session.
    Trả về True nếu thực sự recover, False nếu skip (vừa recover xong)."""
    global _http_session, _last_recover_time
    if not _recover_lock.acquire(blocking=True, timeout=120):
        return False
    try:
        if time.time() - _last_recover_time < 20:
            return False  # thread khác vừa recover, bỏ qua
        log.warning(f"  [Recover] {reason}")
        if _STOP.is_set():
            return False
        # Thử tối đa 3 lần, mỗi lần verify session bằng test request
        for i in range(3):
            _chrome_load_home()
            with _session_lock:
                _http_session = None
                ok = init_http_session()
            if not ok:
                time.sleep(5)
                continue
            # Verify: test request thực tế
            try:
                proxies = {"http": LOCAL_PROXY, "https": LOCAL_PROXY} if USE_PROXY else None
                r = _http_session.get(BASE_URL, timeout=15, proxies=proxies)
                if r.status_code == 200 and "Just a moment" not in r.text:
                    _last_recover_time = time.time()
                    log.info(f"  [Recover] Session verified OK (lan {i+1}).")
                    return True
                log.warning(f"  [Recover] Verify that bai (HTTP {r.status_code}), thu lai...")
            except Exception as e:
                log.warning(f"  [Recover] Verify error: {e}, thu lai...")
            time.sleep(random.uniform(5, 10))
        # Hết lần thử — cập nhật time để tránh spam
        _last_recover_time = time.time()
        log.warning("  [Recover] Het lan thu, tiep tuc voi session hien tai.")
        return True
    finally:
        _recover_lock.release()

def _throttle():
    global _request_count
    with _count_lock:
        _request_count += 1
        cnt = _request_count
    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
    if cnt % COOKIE_REFRESH == 0:
        log.info(f"  [Cookie refresh] Sau {cnt} request...")
        with _session_lock:
            init_http_session()
    if cnt % BURST_EVERY == 0:
        rest = random.uniform(BURST_REST_MIN, BURST_REST_MAX)
        log.info(f"  [Burst rest] {rest:.1f}s sau {cnt} request...")
        time.sleep(rest)

def fetch_html_fast(url: str) -> str | None:
    global _http_session
    if _http_session is None:
        with _session_lock:
            if _http_session is None and not init_http_session():
                return None

    attempt = 0
    while not _STOP.is_set():
        if attempt >= FAIL_BEFORE_RECOVER:
            did = _do_recover(f"HTTP that bai {attempt} lan lien tiep", url=url)
            attempt = 0
            time.sleep(random.uniform(3, 6) if did else random.uniform(20, 35))
            continue

        try:
            proxies = {"http": LOCAL_PROXY, "https": LOCAL_PROXY} if USE_PROXY else None
            r = _http_session.get(url, timeout=20, allow_redirects=True, proxies=proxies)

            final_url = str(getattr(r, "url", url))
            if "hsctvn" not in final_url:
                _do_recover(f"Redirect sang trang la: {final_url[:80]}", url=url)
                attempt = 0
                continue

            if r.status_code in (403, 503) or "Just a moment" in r.text:
                did = _do_recover(f"CF block (HTTP {r.status_code})", url=url)
                attempt = 0
                time.sleep(random.uniform(3, 6) if did else random.uniform(20, 35))
                continue

            if r.status_code == 200:
                _throttle()
                return r.text

            log.warning(f"HTTP {r.status_code} [{url[:60]}] — thu lai...")

        except Exception as e:
            err = str(e)
            if "proxy" in err.lower() or "connection to proxy" in err.lower():
                # Proxy drop kết nối (xoay IP) — tạo lại session để clear pool
                log.warning(f"Proxy drop [{url[:60]}] — reset session, retry...")
                with _session_lock:
                    _http_session = None
                    init_http_session()
                time.sleep(random.uniform(2, 4))
                continue  # không tính vào attempt
            log.warning(f"Fast fetch [{url[:60]}]: {e} — lan {attempt + 1}...")

        attempt += 1
        if not _STOP.is_set():
            wait = min(random.uniform(2, 5) * attempt, 30)
            time.sleep(wait)

    return None

def _ip_rotation_watcher():
    """Proactive reload Chrome ngay sau khi IP xoay để có cf_clearance mới."""
    interval = IP_ROTATE_EVERY + 10
    log.info(f"[IP watcher] Bat dau — reload Chrome moi {interval}s")
    while not _STOP.is_set():
        for _ in range(interval):
            if _STOP.is_set():
                return
            time.sleep(1)
        if _STOP.is_set():
            return
        log.info("[IP watcher] IP vua xoay — cap nhat cf_clearance...")
        _do_recover("IP xoay theo lich", url=BASE_URL)

# ─── Parse detail ─────────────────────────────────────────────────────────────

def normalize_phone(phone: str) -> str:
    if not phone:
        return ""
    d = re.sub(r"\D", "", phone)
    if d.startswith("84") and len(d) == 11:
        d = "0" + d[2:]
    return d if len(d) in (9, 10) and d.startswith("0") else ""

def decode_cf_email(encoded: str) -> str:
    try:
        enc = bytes.fromhex(encoded)
        key = enc[0]
        return "".join(chr(b ^ key) for b in enc[1:]).lower().strip()
    except Exception:
        return ""

def parse_detail(html: str) -> dict:
    soup   = BeautifulSoup(html, "html.parser")
    result = {"so_dien_thoai": "", "email": "", "nam_thanh_lap": "", "nganh_nghe": ""}
    scope  = soup.find("div", class_=lambda c: c and "detail" in c) or soup

    for ul in scope.find_all("ul", class_="hsct"):
        for li in ul.find_all("li"):
            li_str = str(li)
            for icon in li.find_all("i"):
                icon.decompose()
            text      = li.get_text(separator=" ", strip=True)
            text_norm = unidecode(text).lower()

            if "fa-phone" in li_str or "dien thoai" in text_norm:
                nums = re.findall(r"0\d{8,9}", text)
                if nums and not result["so_dien_thoai"]:
                    result["so_dien_thoai"] = normalize_phone(nums[0])

            elif "fa-envelope" in li_str or "email" in text_norm:
                cf_tag = li.find(class_="__cf_email__")
                if cf_tag and cf_tag.get("data-cfemail"):
                    result["email"] = decode_cf_email(cf_tag["data-cfemail"])
                else:
                    emails = re.findall(r"[\w.+-]+@[\w-]+\.[\w.]+", text)
                    if emails and not result["email"]:
                        result["email"] = emails[0].lower()

            elif "fa-calendar" in li_str or "ngay cap" in text_norm or "ngay thanh lap" in text_norm:
                years = re.findall(r"\b(?:19|20)\d{2}\b", text)
                if years and not result["nam_thanh_lap"]:
                    result["nam_thanh_lap"] = years[0]

            elif "fa-tags" in li_str or "nganh nghe" in text_norm or "linh vuc" in text_norm:
                if ":" in text and not result["nganh_nghe"]:
                    result["nganh_nghe"] = text.split(":", 1)[1].strip()

    return result

# ─── Enrich ──────────────────────────────────────────────────────────────────

def enrich_records(batch_size: int = 50) -> list:
    if not os.path.exists(INPUT_PATH):
        log.error(f"Khong tim thay: {INPUT_PATH}")
        return []
    records = pd.read_csv(INPUT_PATH, dtype=str).fillna("").to_dict("records")
    log.info(f"Load {len(records)} records tu {INPUT_PATH}")

    if os.path.exists(LISTING_PATH) and not any(r.get("detail_href") for r in records):
        df_list = pd.read_csv(LISTING_PATH, dtype=str).fillna("")
        if "detail_href" in df_list.columns:
            href_map = df_list.set_index("ma_so_thue")["detail_href"].to_dict()
            for r in records:
                r["detail_href"] = href_map.get(r["ma_so_thue"], "")
            found = sum(1 for r in records if r.get("detail_href"))
            log.info(f"  Merge detail_href: {found}/{len(records)}")

    ENRICH_FIELDS = ["so_dien_thoai", "email", "nam_thanh_lap", "nganh_nghe"]
    attempted_mst = set()

    if os.path.exists(CKPT_PATH):
        df_cp    = pd.read_csv(CKPT_PATH, dtype=str).fillna("")
        ckpt_map = df_cp.set_index("ma_so_thue").to_dict("index")
        attempted_mst = set(df_cp["ma_so_thue"].tolist())

        if RETRY_INCOMPLETE:
            incomplete = set(
                df_cp[
                    df_cp["so_dien_thoai"].str.strip().eq("") &
                    df_cp["email"].str.strip().eq("")
                ]["ma_so_thue"].tolist()
            )
            attempted_mst -= incomplete
            log.info(f"RETRY_INCOMPLETE: thu lai {len(incomplete)} record")

        for r in records:
            if r["ma_so_thue"] in ckpt_map:
                for k in ENRICH_FIELDS:
                    v = ckpt_map[r["ma_so_thue"]].get(k, "")
                    if v:
                        r[k] = v
        log.info(f"Checkpoint: da xu ly {len(attempted_mst)} records")

    pending = [r for r in records if r["ma_so_thue"] not in attempted_mst and r.get("detail_href")]
    sdt_ok  = sum(1 for r in records if r.get("so_dien_thoai"))

    if ENRICH_FROM > 1:
        skip = min(ENRICH_FROM - 1, len(pending))
        log.info(f"ENRICH_FROM={ENRICH_FROM}: bo qua {skip} record dau")
        pending = pending[skip:]

    log.info(f"=== Enrich {len(records)} records | can fetch: {len(pending)} | PARALLEL={PARALLEL} ===")

    if not pending:
        return records

    # Warm-up Chrome để lấy CF cookie
    log.info("Warm-up Chrome pass CF...")
    _chrome_load_home()
    init_http_session()

    # Khởi động IP watcher
    if USE_PROXY:
        threading.Thread(target=_ip_rotation_watcher, daemon=True, name="ip-watcher").start()

    save_lock  = Lock()
    done_lock  = Lock()
    done_count = 0

    def _fetch_one(r: dict) -> dict:
        if _STOP.is_set():
            return r
        url  = f"{BASE_URL}/{r['detail_href']}"
        html = fetch_html_fast(url)
        if html:
            for k, v in parse_detail(html).items():
                if v:
                    r[k] = v
        return r

    def _after_each(r: dict):
        nonlocal done_count, sdt_ok
        with done_lock:
            done_count += 1
            if r.get("so_dien_thoai"):
                sdt_ok += 1
            dc = done_count
        if dc % 10 == 0 or dc == len(pending):
            pct = dc / len(pending) * 100
            log.info(f"  Enrich {dc}/{len(pending)} ({pct:.0f}%) | SDT: {sdt_ok}")
        if dc % batch_size == 0 or dc == len(pending):
            with save_lock:
                pd.DataFrame(records).to_csv(CKPT_PATH, index=False, encoding="utf-8-sig")

    with ThreadPoolExecutor(max_workers=PARALLEL) as ex:
        futures = {ex.submit(_fetch_one, r): r for r in pending}
        for fut in as_completed(futures):
            if _STOP.is_set():
                ex.shutdown(wait=False, cancel_futures=True)
                break
            _after_each(fut.result())

    if _STOP.is_set():
        log.info("Dung theo yeu cau. Luu checkpoint...")
        with save_lock:
            pd.DataFrame(records).to_csv(CKPT_PATH, index=False, encoding="utf-8-sig")

    return records


COLS = [
    "ma_so_thue", "ten_cong_ty", "nam_thanh_lap",
    "so_dien_thoai", "email", "dia_chi", "nganh_nghe",
    "tinh_thanh", "detail_href", "crawled_at",
]

def save_result(records: list) -> pd.DataFrame:
    df = pd.DataFrame(records)
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[COLS].drop_duplicates(subset=["ma_so_thue"], keep="last").reset_index(drop=True)

    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Doanh Nghiep")
        ws = writer.sheets["Doanh Nghiep"]
        for col, w in zip("ABCDEFGHI", [15, 45, 8, 13, 30, 55, 30, 10, 18]):
            ws.column_dimensions[col].width = w

    print(f"\nDa luu:\n  {CSV_PATH}\n  {EXCEL_PATH}")
    total = len(df)
    print(f"\nChat luong {total:,} records:")
    for col, label in [
        ("ma_so_thue",    "MST       "),
        ("ten_cong_ty",   "Ten CT    "),
        ("dia_chi",       "Dia chi   "),
        ("nam_thanh_lap", "Nam TL    "),
        ("so_dien_thoai", "SDT       "),
        ("email",         "Email     "),
        ("nganh_nghe",    "Nganh nghe"),
    ]:
        n = df[col].str.strip().str.len().gt(0).sum()
        print(f"  {label}: {n:>6,} ({n/total*100:.0f}%)")
    return df


# ─── Entry point ─────────────────────────────────────────────────────────────

try:
    records = enrich_records(batch_size=50)
    if records:
        save_result(records)
    else:
        log.warning("Khong co record nao de luu.")
finally:
    close_browser()
    if os.path.exists(CKPT_PATH):
        try:
            saved = pd.read_csv(CKPT_PATH, dtype=str).fillna("").to_dict("records")
            if saved:
                log.info(f"[Finally] Xuat ket qua tu checkpoint ({len(saved)} records)...")
                save_result(saved)
        except Exception as e:
            log.error(f"[Finally] Khong xuat duoc: {e}")
