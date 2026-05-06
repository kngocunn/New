"""
crawler_v3.py — dùng DrissionPage (kết nối Chrome thật, vượt CF tốt hơn Playwright)
Cài: pip install DrissionPage
Chạy: python crawler_v3.py
"""

import re, os, sys, time, random, logging
from datetime import datetime
from unidecode import unidecode
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
CKPT_PATH  = "data/checkpoint.csv"

CHAY_HA_NOI = False 
CHAY_HCM    = False
TARGET      = 500
CO_ENRICH   = True

# Trang bắt đầu cho mỗi tỉnh (1 = từ đầu)
START_PAGE  = {"ha-noi": 1, "ho-chi-minh": 1}
LISTING_PATH = "data/listing.csv"   # lưu kết quả listing trước khi enrich

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
    # headless=False (mặc định) để user giải CF nếu cần
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
    """
    br = get_browser()

    try:
        br.get(url)
    except Exception as e:
        log.error(f"Loi navigate [{url}]: {e}")
        return None

    t0      = time.time()
    warned  = False

    while True:
        # Kiểm tra selector đã xuất hiện chưa
        el = br.ele(f"css:{wait_selector}", timeout=0)
        if el:
            html = br.html
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
            return html

        elapsed = time.time() - t0
        if elapsed > timeout:
            log.error(f"Timeout {timeout}s: [{url}] — khong tim thay {wait_selector!r}")
            return None

        # Cảnh báo nếu chờ lâu (có thể đang bị CF)
        if not warned and elapsed > 8:
            title = ""
            try:
                title = br.title
            except Exception:
                pass
            is_cf = any(s in title for s in ("Just a moment", "Verify", "Checking"))
            print("\n" + "=" * 60)
            print(f"  Chua co du lieu sau {elapsed:.0f}s  |  title: {title!r}")
            if is_cf:
                print("  CLOUDFLARE CHALLENGE — hay giai xac thuc trong Chrome")
            else:
                print("  Trang chua tai xong hoac bi chan — kiem tra cua so Chrome")
            print("  Script se TU DONG TIEP TUC sau khi du lieu xuat hien")
            print("=" * 60 + "\n")
            warned = True

        time.sleep(1.5)


# ─── Parse listing ───────────────────────────────────────────────────────────

def normalize_phone(phone: str) -> str:
    if not phone:
        return ""
    d = re.sub(r"\D", "", phone)
    if d.startswith("84") and len(d) == 11:
        d = "0" + d[2:]
    return d if len(d) in (9, 10) and d.startswith("0") else phone.strip()


def normalize_address(raw: str) -> str:
    return " ".join(raw.split()).replace(" ,", ",") if raw else ""


def decode_cf_email(encoded: str) -> str:
    try:
        enc = bytes.fromhex(encoded)
        key = enc[0]
        return "".join(chr(b ^ key) for b in enc[1:]).lower().strip()
    except Exception:
        return ""


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


# ─── Parse detail ─────────────────────────────────────────────────────────────

def parse_detail(html: str) -> dict:
    soup   = BeautifulSoup(html, "html.parser")
    result = {"so_dien_thoai": "", "email": "", "nam_thanh_lap": "", "nganh_nghe": ""}

    scope = soup.find("div", class_=lambda c: c and "detail" in c) or soup

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


# ─── Crawl ───────────────────────────────────────────────────────────────────

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
    page_num    = START_PAGE.get(tinh_key, 1)
    empty_pages = 0

    log.info(f"=== Crawl listing: {tinh_label} | tu trang {page_num} | muc tieu: {max_records} ===")

    while len(all_records) < max_records:
        url  = get_listing_url(tinh_key, page_num)
        html = fetch_html(url, wait_selector="ul.hsdn")
        if not html:
            log.error(f"Khong lay duoc trang {page_num}, dung.")
            break

        records   = parse_listing(html, tinh_label)
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

        if new_count == 0:
            empty_pages += 1
            if empty_pages >= 2:
                log.info("  2 trang trong lien tiep, dung.")
                break
        else:
            empty_pages = 0

        page_num += 1
        if page_num % 10 == 0:
            wait = random.uniform(8, 12)
            log.info(f"  Nghi {wait:.1f}s ...")
            time.sleep(wait)

    log.info(f"Xong listing {tinh_label}: {len(all_records)} records")
    return all_records


def enrich_records(records: list, batch_size: int = 50) -> list:
    # Nếu không có records trong memory, load từ file có sẵn
    if not records:
        for src in [LISTING_PATH, CKPT_PATH, CSV_PATH]:
            if os.path.exists(src):
                records = pd.read_csv(src, dtype=str).fillna("").to_dict("records")
                log.info(f"Load {len(records)} records tu {src}")
                break

    ENRICH_FIELDS = ["so_dien_thoai", "email", "nam_thanh_lap", "nganh_nghe"]
    enriched_mst  = set()
    if os.path.exists(CKPT_PATH):
        df_cp    = pd.read_csv(CKPT_PATH, dtype=str).fillna("")
        ckpt_map = df_cp.set_index("ma_so_thue").to_dict("index")
        # Chỉ bỏ qua record đã có ít nhất 1 trường enrich
        has_data = df_cp[ENRICH_FIELDS].apply(
            lambda row: any(str(v).strip() for v in row), axis=1
        )
        enriched_mst = set(df_cp[has_data]["ma_so_thue"].tolist())
        for r in records:
            if r["ma_so_thue"] in ckpt_map:
                for k in ENRICH_FIELDS:
                    v = ckpt_map[r["ma_so_thue"]].get(k, "")
                    if v:
                        r[k] = v
        log.info(f"Checkpoint: {len(enriched_mst)} da co du lieu, {len(ckpt_map)-len(enriched_mst)} se thu lai")

    total = len(records)
    log.info(f"=== Enrich {total} records ===")

    for i, r in enumerate(records):
        done = i + 1
        if r["ma_so_thue"] in enriched_mst:
            continue

        href = r.get("detail_href", "")
        if href:
            html = fetch_html(f"{BASE_URL}/{href}", wait_selector="ul.hsct")
            if html:
                r.update(parse_detail(html))
        else:
            log.warning(f"Khong co href: {r['ma_so_thue']}")

        if done % 10 == 0 or done == total:
            sdt_ok = sum(1 for x in records[:done] if x.get("so_dien_thoai"))
            log.info(f"  Enrich {done}/{total} ({done/total*100:.0f}%) | SDT: {sdt_ok}")

        if done % batch_size == 0 or done == total:
            pd.DataFrame(records).to_csv(CKPT_PATH, index=False, encoding="utf-8-sig")

    return records


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

    if CO_ENRICH:
        # Enrich dùng file listing (bao gồm cả records từ các lần chạy trước)
        tat_ca = enrich_records(tat_ca, batch_size=50)

    if tat_ca:
        save_and_merge(tat_ca)
    else:
        log.warning("Khong co record nao.")

finally:
    close_browser()
