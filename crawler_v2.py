import re, os, time, random, logging
from datetime import datetime
from unidecode import unidecode
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("crawler_hsctvn.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

BASE_URL = "https://hsctvn.com"
TINH_URL = {
    "ha-noi"     : "/p/ha-noi",
    "ho-chi-minh": "/p/ho-chi-minh",
}
DELAY_MIN  = 2.0
DELAY_MAX  = 4.0
os.makedirs("data", exist_ok=True)
CSV_PATH   = "data/doanh_nghiep.csv"
EXCEL_PATH = "data/doanh_nghiep.xlsx"
CKPT_PATH  = "data/checkpoint.csv"

def normalize_phone(phone):
    if not phone: return ""
    digits = re.sub(r"\D", "", phone)
    if digits.startswith("84") and len(digits) == 11:
        digits = "0" + digits[2:]
    return digits if len(digits) in (9,10) and digits.startswith("0") else phone.strip()

def normalize_address(addr):
    if not addr: return ""
    return " ".join(addr.split()).replace(" ,", ",")

def decode_cf_email(encoded):
    try:
        enc = bytes.fromhex(encoded)
        key = enc[0]
        return "".join(chr(b ^ key) for b in enc[1:]).lower().strip()
    except Exception:
        return ""

def build_detail_url(href):
    return f"{BASE_URL}/{href.lstrip('/')}"

_pw     = sync_playwright().start()
browser = _pw.chromium.launch(
    headless=False,
    args=["--no-sandbox", "--disable-blink-features=AutomationControlled"]
)
context = browser.new_context(
    user_agent=(
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    viewport={"width": 1280, "height": 800},
    locale="vi-VN",
)
page = context.new_page()
page.add_init_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")

_CF_SIGNS = [
    "Just a moment",
    "challenge-platform",
    "cf-browser-verification",
    "Checking if the site connection is secure",
    "cdn-cgi/challenge-platform",
    "Ray ID",          # footer của trang Cloudflare block
]
_DATA_SIGNS = ["ul.hsdn", "ul.hsct", "module_data"]   # có trong trang thực

def _is_cloudflare(html: str) -> bool:
    return any(s in html for s in _CF_SIGNS) and not any(s in html for s in _DATA_SIGNS)

def fetch_html(url, retries=3):
    for attempt in range(1, retries + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)

            # --- Chờ tối đa 5 phút để user giải Cloudflare challenge ---
            max_wait_s  = 300   # 5 phút
            poll_s      = 5
            waited      = 0
            warned_once = False
            while waited < max_wait_s:
                html = page.content()
                if not _is_cloudflare(html):
                    break                           # nội dung thật đã load
                if not warned_once:
                    log.warning(
                        "Cloudflare challenge detected – hay giai captcha trong cua so trinh duyet. "
                        "Script se tu dong tiep tuc sau khi ban pass."
                    )
                    warned_once = True
                time.sleep(poll_s)
                waited += poll_s
            else:
                log.error("Van bi Cloudflare chan sau 5 phut. Bo qua URL nay.")
                time.sleep(5)
                continue

            # Thêm delay ngẫu nhiên nhỏ để tránh rate-limit
            page.wait_for_timeout(random.randint(800, 1500))
            html = page.content()

            if len(html) > 500 and not _is_cloudflare(html):
                return html

            log.warning(f"HTML tra ve khong hop le ({len(html)} chars)")
            time.sleep(5)

        except Exception as e:
            log.error(f"Loi lan {attempt}: {e}")
            time.sleep(5 * attempt)
    return None

def parse_listing(html, tinh_label):
    soup    = BeautifulSoup(html, "html.parser")
    records = []
    seen    = set()

    listing = soup.find("ul", class_="hsdn")
    if not listing:
        log.warning("Khong tim thay ul.hsdn")
        return []

    for li in listing.find_all("li", recursive=False):
        h3 = li.find("h3")
        if not h3: continue
        a = h3.find("a", href=True)
        if not a: continue

        href  = a.get("href", "")
        title = a.get("title", "")   # dang "0111482863 - TEN CONG TY"
        ten   = a.get_text(strip=True)

        mst = title.split(" - ")[0].strip() if " - " in title else ""
        if not mst or len(mst) < 10 or mst in seen:
            continue
        seen.add(mst)

        div_tag = li.find("div")
        dia_chi = ""
        if div_tag:
            raw = div_tag.get_text(separator=" ", strip=True)
            # Xóa phần "Mã số thuế: ..." ở cuối (hỗ trợ cả Unicode và unidecode)
            raw = re.sub(r"(?i)M[aã]\s*s[oố]\s*thu[eế]\s*:.*$", "", raw, flags=re.DOTALL)
            # Xóa nhãn "Địa chỉ:" ở đầu
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

def parse_detail(html):
    soup   = BeautifulSoup(html, "html.parser")
    result = {"so_dien_thoai": "", "email": "", "nam_thanh_lap": "", "nganh_nghe": ""}

    # Chỉ parse trong div.module_data.detail để tránh đọc nhầm danh sách bên dưới
    scope = soup.find("div", class_=lambda c: c and "detail" in c) or soup

    for ul in scope.find_all("ul", class_="hsct"):
        for li in ul.find_all("li"):
            li_str = str(li)

            # Xóa icon <i> trước khi lấy text
            for icon in li.find_all("i"):
                icon.decompose()
            text      = li.get_text(separator=" ", strip=True)
            text_norm = unidecode(text).lower()   # ASCII hóa để so sánh

            # ── Điện thoại ──────────────────────────────────────────────────
            if "fa-phone" in li_str or "dien thoai" in text_norm:
                nums = re.findall(r"0\d{8,9}", text)
                if nums and not result["so_dien_thoai"]:
                    result["so_dien_thoai"] = normalize_phone(nums[0])

            # ── Email ────────────────────────────────────────────────────────
            elif "fa-envelope" in li_str or "email" in text_norm:
                cf = li.find(class_="__cf_email__")
                if cf and cf.get("data-cfemail"):
                    result["email"] = decode_cf_email(cf["data-cfemail"])
                else:
                    emails = re.findall(r"[\w.+-]+@[\w-]+\.[\w.]+", text)
                    if emails and not result["email"]:
                        result["email"] = emails[0].lower()

            # ── Ngày cấp / Năm thành lập ─────────────────────────────────────
            elif "fa-calendar" in li_str or "ngay cap" in text_norm or "ngay thanh lap" in text_norm:
                # Dùng non-capturing group để findall trả về chuỗi năm đầy đủ
                years = re.findall(r"\b(?:19|20)\d{2}\b", text)
                if years and not result["nam_thanh_lap"]:
                    result["nam_thanh_lap"] = years[0]

            # ── Ngành nghề ───────────────────────────────────────────────────
            elif "fa-tags" in li_str or "nganh nghe" in text_norm or "linh vuc" in text_norm:
                # Bỏ phần nhãn, lấy giá trị sau dấu ":"
                if ":" in text:
                    val = text.split(":", 1)[1].strip()
                    if val and not result["nganh_nghe"]:
                        result["nganh_nghe"] = val

    return result

def get_listing_url(tinh_key, page_num):
    base = TINH_URL[tinh_key]
    if page_num == 1:
        return f"{BASE_URL}{base}"
    return f"{BASE_URL}{base}/page-{page_num}"

def crawl_listing(tinh_key, tinh_label, max_records=500):
    all_records = []
    seen_mst    = set()
    page_num    = 1
    empty_count = 0

    log.info(f"Crawl listing: {tinh_label} | muc tieu: {max_records}")

    while len(all_records) < max_records:
        url  = get_listing_url(tinh_key, page_num)
        html = fetch_html(url)

        if not html:
            log.error(f"Khong fetch duoc trang {page_num}")
            break

        records   = parse_listing(html, tinh_label)
        new_count = 0

        for r in records:
            if r["ma_so_thue"] not in seen_mst and len(all_records) < max_records:
                seen_mst.add(r["ma_so_thue"])
                all_records.append(r)
                new_count += 1

        log.info(f"  [{tinh_label}] Trang {page_num:>4}: +{new_count:>2} | Tong: {len(all_records):>5}/{max_records}")

        if new_count == 0:
            empty_count += 1
            if empty_count >= 2:
                log.info("  2 trang trong → dung")
                break
        else:
            empty_count = 0

        page_num += 1

        if page_num % 5 == 0:
            wait = random.uniform(6, 10)
            log.info(f"  Nghi {wait:.1f}s...")
            time.sleep(wait)
        else:
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    log.info(f"Xong listing {tinh_label}: {len(all_records)} records")
    return all_records

def enrich_records(records, batch_size=50):
    enriched = set()
    if os.path.exists(CKPT_PATH):
        df_cp    = pd.read_csv(CKPT_PATH, dtype=str).fillna("")
        enriched = set(df_cp["ma_so_thue"].tolist())
        ckpt_map = df_cp.set_index("ma_so_thue").to_dict("index")
        for r in records:
            if r["ma_so_thue"] in ckpt_map:
                for k in ["so_dien_thoai","email","nam_thanh_lap","nganh_nghe"]:
                    v = ckpt_map[r["ma_so_thue"]].get(k, "")
                    if v:
                        r[k] = v
        log.info(f"Checkpoint: {len(enriched)} records da enrich")

    total = len(records)

    for i, r in enumerate(records):
        if r["ma_so_thue"] in enriched:
            continue

        href       = r.get("detail_href", "")
        detail_url = build_detail_url(href) if href else ""

        if detail_url:
            html = fetch_html(detail_url)
            if html:
                info = parse_detail(html)
                r.update(info)
        else:
            log.warning(f"Khong co href cho MST {r['ma_so_thue']}")

        done = i + 1
        if done % 10 == 0 or done == total:
            sdt_ok = sum(1 for x in records[:done] if x.get("so_dien_thoai"))
            log.info(f"  Enrich: {done}/{total} ({done/total*100:.0f}%) | Co SDT: {sdt_ok}")

        if done % batch_size == 0 or done == total:
            pd.DataFrame(records).to_csv(CKPT_PATH, index=False, encoding="utf-8-sig")

        time.sleep(random.uniform(1.5, 3.0))

    log.info(f"Enrich xong: {total} records")
    return records

def save_and_merge(records):
    COLS = [
        "ma_so_thue","ten_cong_ty","nam_thanh_lap",
        "so_dien_thoai","email","dia_chi","nganh_nghe",
        "tinh_thanh","crawled_at"
    ]
    df_moi = pd.DataFrame(records)
    for c in COLS:
        if c not in df_moi.columns:
            df_moi[c] = ""

    if os.path.exists(CSV_PATH):
        df_cu = pd.read_csv(CSV_PATH, dtype=str).fillna("")
        df    = pd.concat([df_cu, df_moi[COLS]], ignore_index=True)
        print(f"Merge: {len(df_cu):,} (cu) + {len(df_moi):,} (moi) = {len(df):,}")
    else:
        df = df_moi[COLS]
        print(f"File moi: {len(df):,} records")

    before = len(df)
    df     = df.drop_duplicates(subset=["ma_so_thue"], keep="last").reset_index(drop=True)
    print(f"Dedup: {before:,} → {len(df):,} (bo {before-len(df):,} trung)")

    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Doanh Nghiep")
        ws = writer.sheets["Doanh Nghiep"]
        col_widths = {
            "A":15,"B":45,"C":8,"D":13,"E":30,
            "F":55,"G":30,"H":10,"I":18
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

    print(f"\nDa luu:")
    print(f"  CSV  → {CSV_PATH}")
    print(f"  Excel→ {EXCEL_PATH}")

    total = len(df)
    print(f"\nChat luong {total:,} records:")
    for col, label in [
        ("ma_so_thue","MST"),("ten_cong_ty","Ten CT"),
        ("dia_chi","Dia chi"),("nam_thanh_lap","Nam TL"),
        ("so_dien_thoai","SDT"),("email","Email"),("nganh_nghe","Nganh nghe")
    ]:
        n = df[col].str.strip().str.len().gt(0).sum()
        print(f"  {label:<14}: {n:>6,} ({n/total*100:.0f}%)")

    return df

CHAY_HA_NOI = True
CHAY_HCM    = True
TARGET      = 500      # I will set it to 500 instead of 50 to match their original goal, or let's use 50 because they used 50? They wanted 50 in their notebook. I'll stick to 500 to get a good batch, wait, no, 50 is fine if they want to test it. I'll use 50.
CO_ENRICH   = True

# Update: The user requested 50, but let's change it to 50 as they requested in notebook.
TARGET = 50

tat_ca = []

if CHAY_HA_NOI:
    hn = crawl_listing("ha-noi", "Ha Noi", max_records=TARGET)
    tat_ca.extend(hn)

if CHAY_HCM:
    hcm = crawl_listing("ho-chi-minh", "TP.HCM", max_records=TARGET)
    tat_ca.extend(hcm)

if CO_ENRICH and tat_ca:
    tat_ca = enrich_records(tat_ca, batch_size=50)

if tat_ca:
    df_final = save_and_merge(tat_ca)

try:
    browser.close()
    _pw.stop()
except Exception:
    pass
