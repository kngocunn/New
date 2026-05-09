"""
crawler_v1.py — dùng DrissionPage (kết nối Chrome thật, vượt CF tốt hơn Playwright)
Cài: pip install DrissionPage
Chạy: python crawler_v1.py
"""

import re, os, sys, time, random, logging, signal, json
from threading import Lock, Event
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
from DrissionPage import ChromiumPage, ChromiumOptions

import io
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("crawler_hsctvn.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# ─── Config ──────────────────────────────────────────────────────────────────

BASE_URL   = "https://hsctvn.com"
TINH_URL   = {"ha-noi": "/p/ha-noi", "ho-chi-minh": "/p/ho-chi-minh"}
DELAY_MIN  = 1.5
DELAY_MAX  = 3.0
os.makedirs("data", exist_ok=True)
CSV_PATH   = "data/doanh_nghiep.csv"
EXCEL_PATH = "data/doanh_nghiep.xlsx"

CHAY_HA_NOI = True
CHAY_HCM    = True
TARGET      = 15000

# Trang bắt đầu crawl listing (1 = từ đầu / theo checkpoint)
START_HA_NOI = 117
START_HCM    = 1

START_PAGE     = {"ha-noi": START_HA_NOI, "ho-chi-minh": START_HCM}
LISTING_PATH   = "data/listing.csv"
PAGE_CKPT_PATH = "data/page_checkpoint.json"

FREEZE_AFTER = 5  # số lần thất bại liên tiếp trước khi đóng băng chờ đổi IP

# ─── Stop flag ───────────────────────────────────────────────────────────────

_STOP        = Event()
_FROZEN      = Event()
_freeze_lock = Lock()


def _handle_sigint(sig, frame):
    if not _STOP.is_set():
        print("\n" + "=" * 60)
        print("  [!] Nhan Ctrl+C — se dung sau record hien tai...")
        print("  [!] Nhan Ctrl+C lan 2 de THOAT NGAY.")
        print("=" * 60 + "\n")
        _STOP.set()
        _FROZEN.clear()
    else:
        print("\n[!] Buoc dung ngay!")
        sys.exit(1)

signal.signal(signal.SIGINT, _handle_sigint)


def _freeze_and_wait(reason: str = ""):
    """Đóng băng toàn bộ crawl, chờ người dùng đổi IP rồi nhấn Enter."""
    with _freeze_lock:
        if _FROZEN.is_set():
            return
        _FROZEN.set()
        print("\n" + "=" * 60)
        print(f"  [FROZEN] {reason}")
        print("  Crawl da DONG BANG hoan toan.")
        print("  Hay doi IP, sau do nhan ENTER de tiep tuc...")
        print("  (Nhan Ctrl+C de dung han)")
        print("=" * 60)
        try:
            input("  >> Nhan ENTER sau khi doi IP xong: ")
        except (EOFError, KeyboardInterrupt):
            pass
        _FROZEN.clear()


def _wait_if_frozen():
    """Nếu đang đóng băng, block cho đến khi được giải phóng."""
    if _FROZEN.is_set():
        log.info("  [Thread] Dang cho frozen duoc giai phong...")
        while _FROZEN.is_set() and not _STOP.is_set():
            time.sleep(1)

# ─── Browser ─────────────────────────────────────────────────────────────────

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


# ─── Fetch ───────────────────────────────────────────────────────────────────

def fetch_html(url: str, wait_selector: str = "ul.hsdn",
               timeout: int = 120) -> str | None:
    """
    Điều hướng đến url, chờ wait_selector xuất hiện (tối đa timeout giây).
    Nếu CF hiện: in thông báo, chờ user giải, tự động tiếp tục.
    Nếu thất bại: retry vô hạn cho đến khi thành công hoặc _STOP được set.
    """
    br = get_browser()
    attempt = 0

    while not _STOP.is_set():
        _wait_if_frozen()
        if _STOP.is_set():
            return None

        attempt += 1
        if attempt > FREEZE_AFTER:
            _freeze_and_wait(
                f"Chrome that bai {attempt - 1} lan lien tiep — IP co the bi chan hoan toan\n"
                f"  URL: {url[:80]}"
            )
            attempt = 1
            continue

        try:
            br.get(url)
            attempt = 1
        except Exception as e:
            log.error(f"Loi navigate [{url}]: {e}")
            if _STOP.is_set():
                return None
            wait = min(5 * attempt, 30)
            log.info(f"  Thu lai lan {attempt} sau {wait}s (co the dang doi IP)...")
            time.sleep(wait)
            continue

        t0     = time.time()
        warned = False

        while not _STOP.is_set():
            el = br.ele(f"css:{wait_selector}", timeout=0)
            if el:
                html = br.html
                time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
                return html

            elapsed = time.time() - t0
            if elapsed > timeout:
                title = ""
                try:
                    title = br.title
                except Exception:
                    pass
                is_cf = any(s in title for s in ("Just a moment", "Verify", "Checking", "Attention Required"))
                if is_cf:
                    log.warning(f"CF block sau {timeout}s (lan {attempt}): [{url}] — se thu lai...")
                    break
                else:
                    log.info(f"Timeout {timeout}s: [{url}] — khong co {wait_selector!r}, tra ve html de xu ly")
                    html = br.html
                    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
                    return html

            if not warned and elapsed > 8:
                title = ""
                try:
                    title = br.title
                except Exception:
                    pass
                is_cf = any(s in title for s in ("Just a moment", "Verify", "Checking", "Attention Required"))
                print("\n" + "=" * 60)
                print(f"  Chua co du lieu sau {elapsed:.0f}s  |  title: {title!r}")
                if is_cf:
                    print("  CLOUDFLARE CHALLENGE — hay giai xac thuc trong Chrome")
                else:
                    print("  Trang chua tai xong hoac bi chan — kiem tra cua so Chrome")
                print("  Script se TU DONG TIEP TUC sau khi du lieu xuat hien")
                print("  (Nhan Ctrl+C de dung gracefully, Ctrl+C x2 de thoat ngay)")
                print("=" * 60 + "\n")
                warned = True

            time.sleep(1.5)

        if not _STOP.is_set():
            wait = min(random.uniform(5, 10) * attempt, 60)
            log.info(f"  Thu lai lan {attempt + 1} sau {wait:.1f}s... [{url[:60]}]")
            time.sleep(wait)

    return None


# ─── Parse listing ───────────────────────────────────────────────────────────

def normalize_address(raw: str) -> str:
    return " ".join(raw.split()).replace(" ,", ",") if raw else ""


def parse_listing(html: str, tinh_label: str) -> list:
    soup    = BeautifulSoup(html, "html.parser")
    records = []
    seen    = set()

    listing = soup.find("ul", class_="hsdn")
    if not listing:
        log.warning("Khong tim thay ul.hsdn")
        return []

    for li in listing.find_all("li", recursive=False):
        h3 = li.find("h3")
        if not h3:
            continue
        a = h3.find("a", href=True)
        if not a:
            continue

        href  = a.get("href", "").strip("/")
        title = a.get("title", "")
        ten   = a.get_text(strip=True)

        mst = title.split(" - ")[0].strip() if " - " in title else ""
        if not mst or len(mst) < 10 or mst in seen:
            continue
        seen.add(mst)

        dia_chi = ""
        div_tag = li.find("div")
        if div_tag:
            raw = div_tag.get_text(separator=" ", strip=True)
            raw = re.sub(r"(?i)M[aã]\s*s[oố]\s*thu[eế]\s*:.*$", "", raw, flags=re.DOTALL)
            raw = re.sub(r"(?i)^[^\w]*[Đd][^:]*:\s*", "", raw)
            dia_chi = normalize_address(raw)

        records.append({
            "ma_so_thue"   : mst,
            "ten_cong_ty"  : ten,
            "dia_chi"      : dia_chi,
            "tinh_thanh"   : tinh_label,
            "detail_href"  : href,
            "nam_thanh_lap": "",
            "so_dien_thoai": "",
            "email"        : "",
            "nganh_nghe"   : "",
            "crawled_at"   : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })

    return records


# ─── Crawl ───────────────────────────────────────────────────────────────────

def _load_page_ckpt() -> dict:
    if os.path.exists(PAGE_CKPT_PATH):
        try:
            with open(PAGE_CKPT_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_page_ckpt(tinh_key: str, page_num: int):
    data = _load_page_ckpt()
    data[tinh_key] = page_num
    with open(PAGE_CKPT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def _append_listing(new_records: list):
    """Ghi thêm records mới vào LISTING_PATH, dedup theo MST."""
    df_new = pd.DataFrame(new_records)
    if os.path.exists(LISTING_PATH):
        df_old = pd.read_csv(LISTING_PATH, dtype=str).fillna("")
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new
    df.drop_duplicates(subset=["ma_so_thue"], keep="last", inplace=True)
    df.to_csv(LISTING_PATH, index=False, encoding="utf-8-sig")


def get_listing_url(tinh_key: str, page_num: int) -> str:
    base = TINH_URL[tinh_key]
    return f"{BASE_URL}{base}" if page_num == 1 else f"{BASE_URL}{base}/page-{page_num}"


def crawl_listing(tinh_key: str, tinh_label: str, max_records: int = 500) -> list:
    all_records = []
    seen_mst    = set()

    if os.path.exists(LISTING_PATH):
        try:
            df_ex = pd.read_csv(LISTING_PATH, dtype=str).fillna("")
            prev  = df_ex[df_ex["tinh_thanh"] == tinh_label]
            seen_mst    = set(prev["ma_so_thue"].tolist())
            all_records = prev.to_dict("records")
            if seen_mst:
                log.info(f"  Resume listing {tinh_label}: da co {len(seen_mst)} MST tu file cu")
        except Exception as e:
            log.warning(f"  Khong load duoc listing.csv: {e}")

    manual_start = START_PAGE.get(tinh_key, 1)
    if manual_start > 1:
        page_num = manual_start
        log.info(f"  Bat dau tu trang {page_num} (do nguoi dung chi dinh)")
    else:
        page_ckpt = _load_page_ckpt()
        if tinh_key in page_ckpt and len(seen_mst) > 0:
            page_num = page_ckpt[tinh_key]
            log.info(f"  Resume tu trang {page_num} (checkpoint tu lan chay truoc)")
        else:
            page_num = 1

    log.info(f"=== Crawl listing: {tinh_label} | tu trang {page_num} | muc tieu: {max_records} ===")

    while len(all_records) < max_records and not _STOP.is_set():
        url  = get_listing_url(tinh_key, page_num)
        html = fetch_html(url, wait_selector="ul.hsdn")
        if not html:
            log.info(f"Dung crawl listing theo yeu cau nguoi dung.")
            break

        records = parse_listing(html, tinh_label)

        if not records and "hsdn" not in html:
            log.warning(f"  [{tinh_label}] trang {page_num}: HTML khong co listing, co the load do — thu lai...")
            time.sleep(random.uniform(5, 10))
            continue

        new_count = 0
        new_batch = []
        for r in records:
            if r["ma_so_thue"] not in seen_mst and len(all_records) < max_records:
                seen_mst.add(r["ma_so_thue"])
                all_records.append(r)
                new_batch.append(r)
                new_count += 1

        log.info(
            f"  [{tinh_label}] trang {page_num:>4}: +{new_count:>2}"
            f" | tong {len(all_records):>5}/{max_records}"
        )

        if new_count > 0:
            _append_listing(new_batch)

        page_num += 1
        _save_page_ckpt(tinh_key, page_num)

        if page_num % 10 == 0:
            wait = random.uniform(8, 12)
            log.info(f"  Nghi {wait:.1f}s ...")
            time.sleep(wait)

    log.info(f"Xong listing {tinh_label}: {len(all_records)} records")
    return all_records


COLS = [
    "ma_so_thue", "ten_cong_ty", "nam_thanh_lap",
    "so_dien_thoai", "email", "dia_chi", "nganh_nghe",
    "tinh_thanh", "crawled_at",
]


def save_and_merge(records: list) -> pd.DataFrame:
    df_moi = pd.DataFrame(records)
    for c in COLS:
        if c not in df_moi.columns:
            df_moi[c] = ""

    if os.path.exists(CSV_PATH):
        df_cu = pd.read_csv(CSV_PATH, dtype=str).fillna("")
        df    = pd.concat([df_cu, df_moi[COLS]], ignore_index=True)
        print(f"Merge: {len(df_cu):,} (cu) + {len(df_moi):,} (moi) = {len(df):,}")
    else:
        df = df_moi[COLS].copy()
        print(f"File moi: {len(df):,} records")

    before = len(df)
    df = df.drop_duplicates(subset=["ma_so_thue"], keep="last").reset_index(drop=True)
    print(f"Dedup: {before:,} -> {len(df):,} (bo {before - len(df):,} trung)")

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
    tat_ca = []

    if CHAY_HA_NOI:
        tat_ca.extend(crawl_listing("ha-noi",      "Ha Noi", max_records=TARGET))
    if CHAY_HCM:
        tat_ca.extend(crawl_listing("ho-chi-minh", "TP.HCM", max_records=TARGET))

    if tat_ca:
        save_and_merge(tat_ca)
    else:
        log.warning("Khong co record nao.")

finally:
    close_browser()
