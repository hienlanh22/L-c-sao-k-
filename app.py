# -*- coding: utf-8 -*-
"""
CONG CU TRICH XUAT SAO KE NGAN HANG
- PDF: VietinBank, BIDV, LPBank (7 cot & 10 cot), ACB, MBBank
- HTML: VPBank (T24 Core)
- Excel/CSV: Sacombank, Agribank IPCAS, Excel khac
- Dinh dang ngay chuan: dd/mm/yyyy
- Quan ly Cong ty, So TK, Chu cong ty (1111), Danh ba KH/NCC tach biet tung cong ty.
- Ho tro Hoc quy luat (sua 1 lan, lan sau tu ap dung) + Dong bo Cloud (GitHub / Firebase).
"""
import base64
import csv
from html.parser import HTMLParser
import io
import json
import os
import re
import tempfile
import unicodedata
import urllib.request
import webbrowser
from threading import Timer

from flask import Flask, jsonify, render_template_string, request, send_file
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pdfplumber
import xlrd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


# =========================================================================
#  0. DU LIEU MAC DINH: 16 CONG TY & 29 TAI KHOAN NGAN HANG + 39 KHACH HANG MAU
# =========================================================================
_DEFAULT_ACCOUNTS = [
    ("CÔNG TY TNHH THƯƠNG MẠI KHÔI HƯƠNG", "khoihuong", "LPBank", "8181222666"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ XÂY DỰNG THỊNH HIỀN", "thinhhien", "LPBank", "5565568868"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ XÂY DỰNG THỊNH HIỀN", "thinhhien", "ACB", "77688866"),
    ("CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI XUẤT NHẬP KHẨU HÙNG LONG", "hunglong", "LPBank", "072710420001"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ SẢN XUẤT MINH HUY HMA", "minhhuy", "LPBank", "8386556789"),
    ("CÔNG TY TNHH DV THƯƠNG MẠI PHÚC THÀNH", "phucthanh", "Vietcombank", "34643548S01"),
    ("CÔNG TY TNHH DV THƯƠNG MẠI PHÚC THÀNH", "phucthanh", "MSB", "83988888"),
    ("CÔNG TY TNHH NANO HOÀNG ANH", "hoanganh", "LPBank", "629899999"),
    ("CÔNG TY TNHH DỊCH VỤ THƯƠNG MẠI HÀ KỀN", "haken", "Vietcombank", "1066158986"),
    ("CÔNG TY TNHH XĂNG DẦU THÀNH ANH", "thanhanh", "Vietcombank", "3333206868"),
    ("CÔNG TY TNHH XĂNG DẦU THÀNH ANH", "thanhanh", "Vietinbank", "116000149017"),
    ("CÔNG TY TNHH TMDV LỘC HOÀNG", "lochoang", "BIDV", "8600065578"),
    ("CÔNG TY TRÁCH NHIỆM HỮU HẠN LÂN HIỀN", "lanhien", "BIDV", "4820180367"),
    ("CÔNG TY TRÁCH NHIỆM HỮU HẠN LÂN HIỀN", "lanhien", "Vietinbank", "111000073812"),
    ("CÔNG TY TRÁCH NHIỆM HỮU HẠN LÂN HIỀN", "lanhien", "Agribank", "2902201001656"),
    ("CÔNG TY CỔ PHẦN CƠ KHÍ NÔNG NGHIỆP LÝ NHÂN", "cokhi", "Vietinbank", "116000018872"),
    ("CÔNG TY TNHH TMDV GẠCH NGÓI THUỶ KHANH", "thuykhanh", "Vietinbank", "118003052315"),
    ("CÔNG TY TNHH THƯƠNG MẠI & DỊCH VỤ HOÀNG NGUYÊN", "hoangnguyen", "LPBank", "5115123456"),
    ("CÔNG TY TNHH THƯƠNG MẠI LÊ TRUNG 68", "letrung", "ACB", "98198688"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ VẬN TẢI MINH QUANG", "minhquang", "MB", "1168333999"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ VẬN TẢI MINH QUANG", "minhquang", "Sacombank", "030099253378"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "MB", "2268333999"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "Agribank", "2905201005370"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "NCB", "107909729202"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "LPBank", "077079990001"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "LPBank", "077079990002"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "Sacombank", "030052240022"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "Sacombank", "030052280040"),
    ("CÔNG TY TNHH MỘT THÀNH VIÊN THÀNH PHÁT LÝ NHÂN", "thanhphat", "Co-op Bank", "3500000040867012"),
]

_DEFAULT_LANHIEN_CUSTOMERS = [
    {"code": "0700769778", "name": "CÔNG TY TNHH TƯ VẤN XÂY DỰNG ĐỨC THÀNH", "mst": "0700769778"},
    {"code": "688", "name": "CÔNG TY TNHH AN DƯƠNG 688", "mst": "0700861847"},
    {"code": "88 HÀ NAM", "name": "CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG VÀ THƯƠNG MẠI 88 HÀ NAM", "mst": "0700791847"},
    {"code": "anphu", "name": "CÔNG TY TNHH XÂY DỰNG VÀ TM AN PHÚ", "mst": "0700918003"},
    {"code": "an phú", "name": "CÔNG TY TNHH ĐẦU TƯ VÀ XÂY DỰNG AN PHÚ", "mst": "0700362301"},
    {"code": "binhnghia", "name": "UBND xã Bình Nghĩa", "mst": ""},
    {"code": "CNS", "name": "CÔNG TY TNHH GIẢI PHÁP CÔNG NGHỆ VÀ ĐẦU TƯ XÂY DỰNG CNS", "mst": "0108294154"},
    {"code": "cuongkhoi", "name": "CÔNG TY TNHH THƯƠNG MẠI CƯỜNG KHÔI", "mst": "0700318528"},
    {"code": "DABACO", "name": "CÔNG TY TNHH LỢN GIỐNG DABACO HÀ NAM", "mst": "0700757934"},
    {"code": "DG", "name": "Công Ty TNHH ĐG Trung Kiên", "mst": "0700617302"},
    {"code": "ducthien", "name": "CÔNG TY TNHH TƯ VẤN VÀ ĐẦU TƯ XÂY DỰNG ĐỨC THIỆN", "mst": "0700614397"},
    {"code": "gialinh", "name": "CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG GIA LINH", "mst": "0700775080"},
    {"code": "GIANG PHÚC", "name": "CÔNG TY TNHH XÂY DỰNG VÀ THƯƠNG MẠI GIANG PHÚC", "mst": "0700880494"},
    {"code": "hainam", "name": "CÔNG TY CỔ PHẦN TƯ VẤN ĐTXD HẢI NAM", "mst": "0600447544"},
    {"code": "HDT", "name": "Công ty Cổ phần tập đoàn đầu tư xây dựng và phát triển nhà HDT", "mst": "0101585316"},
    {"code": "hkdnguyenvananh", "name": "HỘ KINH DOANH NGUYỄN VĂN ANH - 1999", "mst": "036099011164"},
    {"code": "hoangha", "name": "CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG VÀ DVTM HOÀNG HÀ", "mst": "2701003690"},
    {"code": "hongphu", "name": "Công Ty TNHH Sản Xuất Thương Mại Và Dịch Vụ Hồng Phú", "mst": "0700776292"},
    {"code": "hungbao", "name": "CÔNG TY TNHH XÂY LẮP HƯNG BẢO", "mst": "0700868698"},
    {"code": "khanhhung", "name": "CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG VÀ DV TM KHÁNH HƯNG", "mst": "0700690334"},
    {"code": "KLKimanh", "name": "Bùi Thị Kim Anh", "mst": ""},
    {"code": "KLTHuy", "name": "Trần Thị Thanh Thủy", "mst": ""},
    {"code": "lam", "name": "Trần Đình Lâm", "mst": ""},
    {"code": "Minh Huyến", "name": "Công ty TNHH đầu tư xây dựng Minh Huyến", "mst": "0700523492"},
    {"code": "minhluat", "name": "Công ty TNHH Minh Luật", "mst": "0700245125"},
    {"code": "Nguyễn Ngọc Trịnh", "name": "Nguyễn Ngọc Trịnh", "mst": "036080018072"},
    {"code": "phuchung", "name": "Công ty Cổ phần đầu tư xây dựng Phục Hưng Capital", "mst": "2700822496"},
    {"code": "phuonghung", "name": "Công Ty CP Đầu Tư Xây Dựng Phượng Hùng", "mst": "0700244146"},
    {"code": "quynhlong", "name": "CÔNG TY CỔ PHẦN QUỲNH LONG", "mst": "0600741257"},
    {"code": "THÀNH HUY", "name": "CÔNG TY CP TẬP ĐOÀN THÀNH HUY", "mst": "3000435260"},
    {"code": "THANH SƠN", "name": "CÔNG TY CỔ PHẦN THƯƠNG MẠI VÀ ĐẦU TƯ THANH SƠN", "mst": "0700432397"},
    {"code": "tmhanam", "name": "CÔNG TY CP XÂY DỰNG DỊCH VỤ THƯƠNG MẠI HÀ NAM", "mst": "0700217336"},
    {"code": "truonghunghn", "name": "CÔNG TY TNHH TRƯỜNG HƯNG HN", "mst": "0700898124"},
    {"code": "tuanthuHN", "name": "CÔNG TY TNHH TUẤN THU HN", "mst": "0700847377"},
    {"code": "TÙNG ANH", "name": "CÔNG TY TNHH CẤP THOÁT NƯỚC TÙNG ANH", "mst": "0700256021"},
    {"code": "vuongthin", "name": "CÔNG TY TNHH VƯƠNG THÌN", "mst": "0700921077"},
    {"code": "XÂY DỰNG 3A", "name": "CÔNG TY TNHH XÂY DỰNG 3A", "mst": "0700805458"},
    {"code": "xuankhanhHN", "name": "CÔNG TY TNHH XUÂN KHÁNH HÀ NAM", "mst": "0700859446"},
    {"code": "xuantung", "name": "Xuân Tùng", "mst": ""},
]

def _build_initial_db():
    companies = {}
    for name, code, bank, acc in _DEFAULT_ACCOUNTS:
        if code not in companies:
            companies[code] = {
                "code": code, "name": name, "accounts": [],
                "owners": [], "customers": [], "suppliers": [], "rules": []
            }
        companies[code]["accounts"].append({"bank": bank, "account": acc})
    if "lanhien" in companies:
        companies["lanhien"]["customers"] = list(_DEFAULT_LANHIEN_CUSTOMERS)
        companies["lanhien"]["owners"] = ["NGUYEN VAN LAN"]
    if "thuykhanh" in companies:
        companies["thuykhanh"]["owners"] = ["TRAN VAN TRUONG", "TRAN THI THUY"]
    return {"companies": companies}

_LOCAL_DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "partners_db.json")

def load_shared_db():
    db = _build_initial_db()
    remote_data = None

    fb_url = os.environ.get("FIREBASE_DB_URL", "").strip().rstrip("/")
    gh_token = os.environ.get("GITHUB_TOKEN", "").strip()
    gh_repo = os.environ.get("GITHUB_REPO", "").strip()

    if fb_url:
        try:
            req = urllib.request.Request(f"{fb_url}/partners_db.json", method="GET")
            with urllib.request.urlopen(req, timeout=5) as resp:
                txt = resp.read().decode("utf-8")
                if txt and txt != "null":
                    remote_data = json.loads(txt)
        except Exception:
            pass
    elif gh_token and gh_repo:
        try:
            url = f"https://api.github.com/repos/{gh_repo}/contents/partners_db.json"
            req = urllib.request.Request(url, headers={
                "Authorization": f"Bearer {gh_token}",
                "Accept": "application/vnd.github+json",
                "User-Agent": "SaoKeApp"
            })
            with urllib.request.urlopen(req, timeout=5) as resp:
                meta = json.loads(resp.read().decode("utf-8"))
                remote_data = json.loads(base64.b64decode(meta["content"]).decode("utf-8"))
        except Exception:
            pass

    if remote_data is None and os.path.exists(_LOCAL_DB_FILE):
        try:
            with open(_LOCAL_DB_FILE, "r", encoding="utf-8") as f:
                remote_data = json.load(f)
        except Exception:
            pass

    if remote_data and isinstance(remote_data.get("companies"), dict):
        for code, cdata in remote_data["companies"].items():
            if code in db["companies"]:
                for k in ("customers", "suppliers", "accounts", "owners", "rules"):
                    if k in cdata:
                        db["companies"][code][k] = cdata[k]
                if cdata.get("name"):
                    db["companies"][code]["name"] = cdata["name"]
            else:
                cdata.setdefault("owners", [])
                cdata.setdefault("rules", [])
                cdata.setdefault("customers", [])
                cdata.setdefault("suppliers", [])
                cdata.setdefault("accounts", [])
                db["companies"][code] = cdata
    return db

def save_shared_db(db):
    try:
        with open(_LOCAL_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    fb_url = os.environ.get("FIREBASE_DB_URL", "").strip().rstrip("/")
    gh_token = os.environ.get("GITHUB_TOKEN", "").strip()
    gh_repo = os.environ.get("GITHUB_REPO", "").strip()

    if fb_url:
        try:
            body = json.dumps(db, ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(
                f"{fb_url}/partners_db.json", data=body, method="PUT",
                headers={"Content-Type": "application/json"}
            )
            urllib.request.urlopen(req, timeout=6)
            return True, "Đã đồng bộ lên Firebase Cloud (mọi máy đều thấy)"
        except Exception as e:
            return False, f"Lỗi lưu Firebase: {e}"

    if gh_token and gh_repo:
        try:
            url = f"https://api.github.com/repos/{gh_repo}/contents/partners_db.json"
            sha = None
            try:
                req_get = urllib.request.Request(url, headers={
                    "Authorization": f"Bearer {gh_token}",
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "SaoKeApp"
                })
                with urllib.request.urlopen(req_get, timeout=5) as resp:
                    sha = json.loads(resp.read().decode("utf-8")).get("sha")
            except Exception:
                pass

            content_b64 = base64.b64encode(json.dumps(db, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
            payload = {"message": "Cap nhat danh ba & quy luat tu Web", "content": content_b64}
            if sha:
                payload["sha"] = sha
            req_put = urllib.request.Request(
                url, data=json.dumps(payload).encode("utf-8"), method="PUT",
                headers={
                    "Authorization": f"Bearer {gh_token}",
                    "Accept": "application/vnd.github+json",
                    "Content-Type": "application/json",
                    "User-Agent": "SaoKeApp"
                }
            )
            urllib.request.urlopen(req_put, timeout=8)
            return True, "Đã đồng bộ lên GitHub (mọi máy đều thấy)"
        except Exception as e:
            return False, f"Lỗi lưu GitHub: {e}"

    return True, "Đã lưu trên máy chủ Web (thêm GITHUB_TOKEN vào Render để đồng bộ vĩnh viễn)"


# =========================================================================
#  1. CHUAN HOA NGAY (dd/mm/yyyy), BO DAU TIENG VIET & TIEN TE
# =========================================================================
_ENG_MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
}

def strip_accents(s):
    if not s:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    return unicodedata.normalize("NFC", s)

def normalize_date_ddmmyyyy(val):
    if not val:
        return ""
    s = str(val).strip()

    m_eng = re.search(r'(?<!\d)(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2,4})(?!\d)', s)
    if m_eng:
        mon = _ENG_MONTHS.get(m_eng.group(2).upper())
        if mon:
            d, y = int(m_eng.group(1)), int(m_eng.group(3))
            y = y + 2000 if y < 100 else y
            return f"{d:02d}/{mon:02d}/{y:04d}"

    m_iso = re.search(r'(?<!\d)(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?!\d)', s)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{m:02d}/{y:04d}"

    m_num4 = re.search(r'(?<!\d)(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?!\d)', s)
    if m_num4:
        d, m, y = int(m_num4.group(1)), int(m_num4.group(2)), int(m_num4.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{m:02d}/{y:04d}"

    m_num2 = re.search(r'(?<!\d)(\d{1,2})[-/](\d{1,2})[-/](\d{2})(?!\d)', s)
    if m_num2:
        d, m, y = int(m_num2.group(1)), int(m_num2.group(2)), int(m_num2.group(3)) + 2000
        if 1 <= m <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{m:02d}/{y:04d}"

    return s


def to_num(s):
    if s is None or s == "" or s == "-NIL-":
        return 0.0
    cleaned = re.sub(r'[^\d.\-]', '', str(s))
    if not cleaned or cleaned in ('-', '.', '-.'):
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _is_money_token(val):
    s = str(val or "").strip()
    if not s or s == "-NIL-":
        return False
    if re.search(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', s):
        return False
    return bool(re.match(r'^-?\d{1,3}([.,]\d{3})*([.,]\d{1,2})?$|^-?\d+([.,]\d{1,2})?$', s))


def _vn_money_to_std_str(val):
    s = str(val or "").strip()
    if not _is_money_token(s):
        return "0"
    if re.match(r'^-?\d{1,3}(\.\d{3})+(,\d+)?$', s):
        s = s.replace(".", "").replace(",", ".")
        try:
            num = abs(float(s))
            return f"{num:,.0f}" if num != 0 else "0"
        except ValueError:
            return "0"
    num = to_num(s.replace(",", ""))
    return f"{abs(num):,.0f}" if num != 0 else "0"


# =========================================================================
#  2. NHAN DIEN NGAN HANG TU NOI DUNG FILE PDF
# =========================================================================
def detect_bank(text):
    t = text or ""
    if "www.acb.com.vn" in t or ("ASIA COMMERCIAL BANK" in t and "Tổng tiền rút ra" in t):
        return "acb"
    if "Ngân Hàng TMCP Công Thương Việt Nam" in t or "eFAST" in t or "vietinbank.vn" in t:
        return "vietinbank"
    if "NGÂN HÀNG TMCP ĐẦU TƯ VÀ PHÁT TRIỂN VIỆT NAM" in t or "Bank for Investment and Development" in t:
        return "bidv"
    if "LPBank" in t or "Loc Phat Viet Nam" in t or "Lộc Phát Việt Nam" in t:
        return "lpbank"
    if "NGÂN HÀNG TMCP QUÂN ĐỘI" in t or "Military Commercial Joint Stock Bank" in t or "MBBank's System" in t:
        return "mbbank"
    return "unknown"


# =========================================================================
#  3A. VIETINBANK PDF
# =========================================================================
def _classify_vtb(w):
    x0, x1 = w['x0'], w['x1']
    if x1 <= 42: return "STT"
    if x0 < 92: return "Ngay"
    if x0 < 255 and x1 < 260: return "MoTa"
    if x1 <= 306: return "No"
    if x1 <= 370: return "Co"
    if x1 <= 436: return "SoDu"
    if x0 < 522: return "DoiTac"
    if x0 < 610: return "SoGD"
    if x0 < 700: return "MaDinhDanh"
    return "NgayPhatSinh"


def extract_vietinbank(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    for page in pdf.pages:
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        ptext = page.extract_text() or ""

        for line in ptext.split("\n"):
            if "Số dư đầu kỳ" in line:
                m = re.search(r'([\d,]+)\s*$', line)
                if m: beginning_balance = m.group(1)
            if "Dear Customer:" in line and not header_info.get('customer'):
                header_info['customer'] = line.split("Dear Customer:", 1)[1].strip()
            if "Tài khoản số:" in line:
                m = re.search(r'Tài khoản số:\s*(\S+)', line)
                if m: header_info['account_no'] = m.group(1)
            if "Doanh số phát sinh" in line:
                nums = re.findall(r'[\d,]+', line)
                if len(nums) >= 2: total_debit, total_credit = nums[-2], nums[-1]
            if "Số dư cuối kỳ" in line:
                m = re.search(r'([\d,]+)\s*$', line)
                if m: ending_balance = m.group(1)

        header_tops = [w['top'] for w in words if w['text'] == 'STT' and w['x0'] < 25]
        table_top = header_tops[0] + 25 if header_tops else 0
        cutoff = [w['top'] for w in words if w['text'] == 'cứu' and w['x0'] < 50]
        table_bottom = min(cutoff) if cutoff else 100000
        end_tops = [w['top'] for w in words if w['text'] == 'kỳ' and any(
            w2['text'] == 'cuối' and abs(w2['top'] - w['top']) < 2 and w2['x0'] < w['x0'] for w2 in words)]
        if end_tops:
            table_bottom = min(table_bottom, min(end_tops) - 3)

        twords = [w for w in words if table_top <= w['top'] < table_bottom]
        stt_anchors = sorted(
            [w for w in twords if _classify_vtb(w) == "STT" and re.match(r'^\d+$', w['text'])],
            key=lambda w: w['top'])
        if not stt_anchors:
            continue

        for idx, anchor in enumerate(stt_anchors):
            top_start = anchor['top'] - 1
            top_end = stt_anchors[idx + 1]['top'] - 1 if idx + 1 < len(stt_anchors) else table_bottom
            row_words = [w for w in twords if top_start <= w['top'] < top_end]
            buckets = {}
            for w in row_words:
                buckets.setdefault(_classify_vtb(w), []).append(w)

            def txt(col):
                ws = sorted(buckets.get(col, []), key=lambda w: (round(w['top']), w['x0']))
                return " ".join(w['text'] for w in ws).strip()

            ngay_words = sorted(buckets.get("Ngay", []), key=lambda w: (w['x0'], round(w['top'])))
            ngay_chuan = ""
            for nw in ngay_words:
                cand = normalize_date_ddmmyyyy(nw['text'])
                if re.match(r'^\d{2}/\d{2}/\d{4}$', cand):
                    ngay_chuan = cand
                    break
            if not ngay_chuan:
                ngay_chuan = normalize_date_ddmmyyyy(txt("Ngay") or txt("NgayPhatSinh"))

            all_rows.append({
                'STT': anchor['text'], 'Ngay': ngay_chuan, 'MoTa': txt("MoTa"),
                'No': txt("No"), 'Co': txt("Co"), 'SoDu': txt("SoDu"),
                'DoiTac': txt("DoiTac"), 'MaGD': txt("SoGD"),
            })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3B. BIDV PDF
# =========================================================================
def extract_bidv(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    first_page_text = pdf.pages[0].extract_text() or ""
    for line in first_page_text.split("\n"):
        line = line.strip()
        if line.startswith("Khách hàng/"):
            m = re.search(r'Customer:\s*(.*?)(?:\s+Địa chỉ|$)', line)
            if m: header_info['customer'] = m.group(1).strip()
        elif line.startswith("Tên tài khoản/") and not header_info.get('customer'):
            m = re.search(r'Account name:\s*(.*?)(?:\s+Loại tiền|$)', line)
            if m: header_info['customer'] = m.group(1).strip()
        elif line.startswith("Số tài khoản/"):
            m = re.search(r'Account No:\s*(\S+)', line)
            if m: header_info['account_no'] = m.group(1)

    for page in pdf.pages:
        for table in page.extract_tables():
            for row in table:
                row = [c.strip() if isinstance(c, str) else c for c in row]
                if not row or row[0] is None:
                    continue
                label = (row[0] or "").replace("\n", " ")

                if "STT" in label and "No" in label: continue
                if "Số dư đầu kỳ" in label or "Opening balance" in label:
                    val = next((c for c in row[1:] if c not in (None, "")), None)
                    if val: beginning_balance = val.replace("\n", "")
                    continue
                if "Cộng phát sinh" in label or "Total Amount" in label:
                    nums = [c.replace("\n", "") for c in row[1:] if c not in (None, "")]
                    if len(nums) >= 2: total_debit, total_credit = nums[0], nums[1]
                    continue
                if "Số dư cuối kỳ" in label or "Closing balance" in label:
                    nums = [c.replace("\n", "") for c in row[1:] if c not in (None, "")]
                    if nums: ending_balance = nums[-1] if len(nums) == 1 else nums[0]
                    continue
                if len(row) < 11: continue
                stt = row[0]
                if not stt or not stt.isdigit(): continue

                all_rows.append({
                    'STT': stt,
                    'Ngay': normalize_date_ddmmyyyy((row[1] or "").replace("\n", " ").strip()),
                    'MoTa': (row[10] or "").replace("\n", " ").strip(),
                    'No': (row[4] or "0").replace("\n", ""),
                    'Co': (row[5] or "0").replace("\n", ""),
                    'SoDu': (row[6] or "").replace("\n", ""),
                    'DoiTac': "",
                    'MaGD': (row[3] or "").replace("\n", "").strip(),
                })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3C. LPBANK PDF (7 cot moi & 10 cot cu)
# =========================================================================
def _clean_lpb(v):
    if not isinstance(v, str) or not v:
        return ""
    v = re.sub(r'^[A-Z0-9]{2,8}\n', '', v)
    return re.sub(r'\s+', ' ', v.replace("\n", " ")).strip()

_DATE_RE_LPB = re.compile(r'^\d{2}/\d{2}/\d{4}$')

def extract_lpbank(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    first_page_text = pdf.pages[0].extract_text() or ""
    for line in first_page_text.split("\n"):
        line = line.strip()
        if line.startswith("Tên tài khoản:") and not header_info.get('customer'):
            m = re.search(r'Tên tài khoản:\s*(.*?)(?:\s+Mã khách hàng|$)', line)
            if m: header_info['customer'] = m.group(1).strip()
        elif line.startswith("Tên khách hàng:"):
            header_info['customer'] = line.split(":", 1)[1].strip()
        elif line.startswith("Số tài khoản:"):
            m = re.search(r'Số tài khoản:\s*(\S+)', line)
            if m: header_info['account_no'] = m.group(1)

    stt_counter = 0
    for page in pdf.pages:
        for table in page.extract_tables():
            for row in table:
                row = [_clean_lpb(c) for c in row]
                if len(row) < 7:
                    continue
                col0 = row[0]
                row_joined = " ".join(row[:4])

                if "Ngày giao dịch" in col0 or "Txn. date" in col0: continue
                if "Cộng doanh số" in row_joined or "Total" in row_joined:
                    nums = [c for c in row[4:] if c and re.search(r'[\d,]+', c)]
                    if len(nums) >= 2: total_debit, total_credit = nums[0], nums[1]
                    continue
                if "Số dư cuối" in row_joined or "Closing Balance" in row_joined:
                    ending_balance = next((c for c in reversed(row) if c and re.search(r'\d', c)), None)
                    continue
                if any("Số dư đầu" in c or "Openning Balance" in c or "Opening Balance" in c for c in row):
                    beginning_balance = next((c for c in reversed(row) if c and re.search(r'\d', c)), None)
                    continue

                if not col0 and all_rows and len(row) >= 3 and row[2] and not any(row[4:]):
                    all_rows[-1]['MoTa'] = (all_rows[-1]['MoTa'] + " " + row[2]).strip()
                    continue

                if not _DATE_RE_LPB.match(col0):
                    continue

                stt_counter += 1
                if len(row) >= 10:
                    no_val, co_val, sodu_val = row[7] or "0", row[8] or "0", row[9]
                    doi_tac, ma_gd = row[4], row[6]
                else:
                    no_val, co_val, sodu_val = row[4] or "0", row[5] or "0", row[6]
                    doi_tac, ma_gd = "", row[3]

                all_rows.append({
                    'STT': str(stt_counter), 'Ngay': normalize_date_ddmmyyyy(row[0]), 'MoTa': row[2],
                    'No': no_val, 'Co': co_val, 'SoDu': sodu_val, 'DoiTac': doi_tac, 'MaGD': ma_gd,
                })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3D. ACB PDF (GIU NGUYEN TU TrichXuatSaoKe (4).py)
# =========================================================================
ACB_NGAY_CHINH = "hieu_luc"
_ACB_DEFAULT_EDGES = [36.0, 83.4, 139.6, 164.3, 322.7, 401.5, 480.2, 559.0]
_ACB_DATE_RE = re.compile(r'\d{2}/\d{2}/\d{4}')
_ACB_NUM_RE = r'(-?[\d.,]+)'

def _acb_amount(s):
    s = (s or "").strip().replace(" ", "")
    if not s: return ""
    neg = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.strip("-()")
    if "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
    elif "," in s: s = s.replace(",", "") if re.fullmatch(r'\d{1,3}(,\d{3})+', s) else s.replace(",", ".")
    elif "." in s and re.fullmatch(r'\d{1,3}(\.\d{3})+', s): s = s.replace(".", "")
    return ("-" if neg else "") + s

def _acb_is_header_fill(r):
    c = r.get('non_stroking_color')
    try:
        return c is not None and len(c) == 3 and abs(c[0] - 0.722) < 0.03 and abs(c[1] - 0.914) < 0.03 and abs(c[2] - 1.0) < 0.03
    except TypeError:
        return False

def _acb_layout(page, words):
    groups = {}
    for r in page.rects:
        if _acb_is_header_fill(r): groups.setdefault(round(r['top']), []).append(r)
    if groups:
        grp = max(groups.values(), key=len)
        if len(grp) == 7:
            xs = sorted(r['x0'] for r in grp)
            return xs + [max(r['x1'] for r in grp)], max(r['bottom'] for r in grp)
    ghi = [w['top'] for w in words if w['text'] == 'Ghi' and 330 < (w['x0'] + w['x1']) / 2 < 410]
    return list(_ACB_DEFAULT_EDGES), (min(ghi) + 14 if ghi else 0.0)

def _acb_row_bounds(tops, anchors, thr=11.7):
    tops = sorted(set(tops))
    cuts = []
    for a, b in zip(anchors, anchors[1:]):
        seg = [t for t in tops if a - 1.0 <= t <= b + 1.0]
        if len(seg) >= 2:
            lo, hi = max(zip(seg, seg[1:]), key=lambda p: p[1] - p[0])
            cuts.append((lo + hi) / 2)
        else:
            cuts.append((a + b) / 2)
    top_edge = anchors[0]
    for t in sorted((t for t in tops if t < anchors[0] - 1.0), reverse=True):
        if top_edge - t > thr: break
        top_edge = t
    bot_edge = anchors[-1]
    for t in sorted(t for t in tops if t > anchors[-1] + 1.0):
        if t - bot_edge > thr: break
        bot_edge = t
    edges = [top_edge - 1.0] + cuts + [bot_edge + 1.0]
    return list(zip(edges[:-1], edges[1:]))

def _acb_page_rows(page):
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    edges, table_top = _acb_layout(page, words)
    def col_of(w):
        c = (w['x0'] + w['x1']) / 2
        for i in range(7):
            if edges[i] <= c < edges[i + 1]: return i
        return None
    cand = [(w, col_of(w)) for w in words if w['top'] >= table_top]
    cand = [(w, c) for w, c in cand if c is not None]
    anchors = sorted(w['top'] for w, c in cand if c == 2 and re.fullmatch(r'\d+', w['text']))
    if not anchors: return []
    bounds = _acb_row_bounds([w['top'] for w, _ in cand], anchors)

    rows = []
    for (lo, hi) in bounds:
        cells = {i: [] for i in range(7)}
        for w, c in cand:
            if lo < w['top'] <= hi: cells[c].append(w)
        def txt(i):
            ws = sorted(cells[i], key=lambda w: (round(w['top']), w['x0']))
            return " ".join(w['text'] for w in ws).strip()
        sogd = next((w['text'] for w in cells[2] if re.fullmatch(r'\d+', w['text'])), "")
        m0 = _ACB_DATE_RE.search(txt(0))
        rows.append({
            'hieu_luc': m0.group(0) if m0 else txt(0), 'giao_dich': txt(1), 'so_gd': sogd,
            'noi_dung': txt(3), 'no': _acb_amount(txt(4)), 'co': _acb_amount(txt(5)), 'so_du': _acb_amount(txt(6)),
        })
    return rows

def _acb_header(page):
    info = {}
    text = page.extract_text() or ""
    m = re.search(r'Thời gian sao kê\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', text)
    if m: info['period'] = f"Tu {m.group(1)} den {m.group(2)}"
    for key, pat in (('beginning', r'Số dư đầu\s+'), ('total_debit', r'Tổng tiền rút ra\s+'),
                     ('total_credit', r'Tổng tiền gửi vào\s+'), ('ending', r'Số dư cuối\s+')):
        mm = re.search(pat + _ACB_NUM_RE, text)
        if mm: info[key] = _acb_amount(mm.group(1))
    words = page.extract_words()
    left = [w for w in words if w['x1'] < 145 and 150 < w['top'] < 340]
    line_tops = sorted({round(w['top'], 1) for w in left})
    labels = {}
    for t in line_tops:
        s = " ".join(w['text'] for w in sorted((w for w in left if abs(w['top'] - t) < 1.5), key=lambda w: w['x0']))
        key = ('name' if s.startswith('Chủ') else 'acc' if s.startswith('Số tài khoản') else None)
        if key: labels.setdefault(key, []).append(t)
    if labels:
        centers = {k: sum(v) / len(v) for k, v in labels.items()}
        lo, hi = min(min(v) for v in labels.values()) - 12, max(max(v) for v in labels.values()) + 12
        vals = [w for w in words if 150 <= w['x0'] < 352 and lo <= w['top'] <= hi]
        got = {}
        for t in sorted({round(w['top'], 1) for w in vals}):
            s = " ".join(w['text'] for w in sorted((w for w in vals if abs(w['top'] - t) < 1.5), key=lambda w: w['x0']))
            k = min(centers, key=lambda kk: abs(centers[kk] - t))
            got.setdefault(k, []).append(s)
        info['customer'] = " ".join(got.get('name', []))
        info['account_no'] = " ".join(got.get('acc', []))
    return info

def extract_acb(pdf):
    hdr = _acb_header(pdf.pages[0])
    header_info = {k: hdr[k] for k in ('period', 'customer', 'account_no') if hdr.get(k)}
    all_rows = []
    for page in pdf.pages:
        for r in _acb_page_rows(page):
            stt = len(all_rows) + 1
            ngay_chinh = r['hieu_luc'] if ACB_NGAY_CHINH == "hieu_luc" else r['giao_dich']
            all_rows.append({
                'STT': str(stt), 'Ngay': normalize_date_ddmmyyyy(ngay_chinh), 'MoTa': r['noi_dung'],
                'No': r['no'] or "0", 'Co': r['co'] or "0", 'SoDu': r['so_du'],
                'DoiTac': "", 'MaGD': r['so_gd'],
            })
    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': hdr.get('beginning'), 'ending_balance': hdr.get('ending'),
        'total_debit': hdr.get('total_debit'), 'total_credit': hdr.get('total_credit'),
    }


# =========================================================================
#  3E. MBBANK PDF
# =========================================================================
def extract_mbbank(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    p0_text = pdf.pages[0].extract_text() or ""
    for line in p0_text.split("\n"):
        line = line.strip()
        if "Customer name:" in line:
            header_info['customer'] = line.split("Customer name:", 1)[1].strip()
        elif "Account No:" in line:
            m = re.search(r'Account No:\s*(\S+)', line)
            if m: header_info['account_no'] = m.group(1)
        elif "Opening Balance:" in line:
            m = re.search(r'Opening Balance:\s*([\d,]+)', line)
            if m: beginning_balance = m.group(1)

    last_text = pdf.pages[-1].extract_text() or ""
    for line in last_text.split("\n"):
        if "Closing Balance:" in line:
            m = re.search(r'Closing Balance:\s*([\d,]+)', line)
            if m: ending_balance = m.group(1)

    for page in pdf.pages:
        for table in page.extract_tables():
            for row in table:
                row = [re.sub(r'\s+', ' ', (c or "").replace("\n", " ")).strip() for c in row]
                if len(row) < 9:
                    continue
                if "TỔNG TIỀN" in row[1] or "TỔNG TIỀN" in row[2]:
                    total_debit, total_credit = row[3] or None, row[4] or None
                    continue
                if not row[0] or not row[0].isdigit():
                    continue
                all_rows.append({
                    'STT': row[0], 'Ngay': normalize_date_ddmmyyyy(row[1] or row[2]), 'MoTa': row[7],
                    'No': row[3] or "0", 'Co': row[4] or "0", 'SoDu': row[5], 'DoiTac': row[6], 'MaGD': row[8],
                })

    all_rows.reverse()
    for i, r in enumerate(all_rows, start=1):
        r['STT'] = str(i)

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3F. VPBANK HTML
# =========================================================================
_VPB_DATE_RE = re.compile(r'^(\d{1,2}) ([A-Za-z]{3}) (\d{2,4})$')

class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.title, self.tables, self._stack = "", [], []
        self._row, self._cell, self._in_title = None, None, False
    def _close_cell(self):
        if self._cell is not None and self._row is not None:
            self._row.append(re.sub(r'\s+', ' ', "".join(self._cell).replace("\xa0", " ")).strip())
        self._cell = None
    def _close_row(self):
        self._close_cell()
        if self._row is not None and self._stack: self._stack[-1].append(self._row)
        self._row = None
    def handle_starttag(self, tag, attrs):
        if tag == "title": self._in_title = True
        elif tag == "table": self._close_row(); self._stack.append([])
        elif tag == "tr": self._close_row(); self._row = []
        elif tag in ("td", "th"):
            self._close_cell()
            if self._row is None: self._row = []
            self._cell = []
    def handle_endtag(self, tag):
        if tag == "title": self._in_title = False
        elif tag in ("td", "th"): self._close_cell()
        elif tag in ("tr", "thead", "tbody", "tfoot"): self._close_row()
        elif tag == "table":
            self._close_row()
            if self._stack: self.tables.append(self._stack.pop())
    def handle_data(self, data):
        if self._in_title: self.title += data
        if self._cell is not None: self._cell.append(data)

def _read_text_file(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1258", "latin-1"):
        try: return raw.decode(enc)
        except UnicodeDecodeError: continue
    return raw.decode("utf-8", errors="replace")

def extract_vpbank_html(parser):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None
    for table in parser.tables:
        for row in table:
            if not row: continue
            label = row[0].upper()
            if label == "NR." or (len(row) >= 3 and row[0].upper().startswith("STATEMENT")):
                header_info['account_no'] = row[-1]
            elif label == "CUSTOMER":
                header_info['customer'] = row[2] if len(row) > 2 else ""
            elif label.startswith("BALANCE AT PERIOD START") and len(row) > 1:
                beginning_balance = row[1]

    data_table, col, header_idx = None, {}, None
    for table in parser.tables:
        for i, row in enumerate(table):
            names = [c.strip().upper() for c in row]
            if "VALDESC" in names and "PDESC" in names:
                data_table, header_idx = table, i
                for j, nme in enumerate(names):
                    if nme and nme not in col: col[nme] = j
                break
        if data_table is not None: break
    if data_table is None:
        raise ValueError("Khong tim thay bang giao dich VPBank")

    def get(row, name):
        j = col.get(name)
        return row[j] if j is not None and j < len(row) else ""

    stt = 0
    for row in data_table[header_idx + 1:]:
        if not row or not any(row): continue
        val_date, pdesc = get(row, "VALDESC"), get(row, "PDESC")
        if _VPB_DATE_RE.match(val_date or ""):
            stt += 1
            all_rows.append({
                'STT': str(stt), 'Ngay': normalize_date_ddmmyyyy(val_date),
                'MoTa': get(row, "NARR") or pdesc,
                'No': (get(row, "PDAMT") or "0").strip().lstrip("-").strip() or "0",
                'Co': (get(row, "PCAMT") or "0").strip().lstrip("-").strip() or "0",
                'SoDu': get(row, "BOOKED BALANCE"), 'DoiTac': "", 'MaGD': pdesc,
            })
        elif "TOTAL MOVEMENT" in (pdesc or "").upper():
            total_debit = (get(row, "PDAMT") or "").strip().lstrip("-").strip() or None
            total_credit = (get(row, "PCAMT") or "").strip().lstrip("-").strip() or None
        elif any(c.upper().startswith("BALANCE AT PERIOD END") for c in row):
            nums = [c for c in row if re.match(r'^-?[\d,]+(\.\d+)?$', c)]
            if nums: ending_balance = nums[0]
    return {'header_info': header_info, 'rows': all_rows, 'beginning_balance': beginning_balance, 'ending_balance': ending_balance, 'total_debit': total_debit, 'total_credit': total_credit}

def extract_html(path):
    try:
        parser = _TableParser()
        parser.feed(_read_text_file(path))
        parser.close()
        res = extract_vpbank_html(parser)
        res['source_file'], res['bank'], res['loi'] = os.path.basename(path), "vpbank", None
        return res
    except Exception as e:
        return {'source_file': os.path.basename(path), 'bank': 'LOI', 'header_info': {}, 'rows': [], 'loi': f"Loi HTML: {e}"}


# =========================================================================
#  3G. EXCEL / CSV (SACOMBANK + AGRIBANK + CHUNG + DANH SACH KH/NCC)
# =========================================================================
def _read_tabular_light(path):
    with open(path, "rb") as f:
        magic = f.read(8)
    if magic.startswith(b'PK'):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        rows = [[("" if c is None else str(c)) for c in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    if magic.startswith(b'\xD0\xCF\x11\xE0'):
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        rows = []
        for r in range(sh.nrows):
            rv = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    rv.append(xlrd.xldate_as_datetime(cell.value, wb.datemode).strftime("%d/%m/%Y"))
                elif cell.ctype == xlrd.XL_CELL_NUMBER and cell.value == int(cell.value):
                    rv.append(str(int(cell.value)))
                else:
                    rv.append(str(cell.value))
            rows.append(rv)
        return rows
    for enc in ("utf-8-sig", "utf-8", "cp1258", "latin-1"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return list(csv.reader(f))
        except Exception: continue
    return []


def parse_partner_excel(path):
    rows = _read_tabular_light(path)
    if not rows:
        raise ValueError("File danh sách trống hoặc không đọc được")

    header_idx = -1
    code_col = name_col = mst_col = -1

    for i, r in enumerate(rows[:15]):
        norm_cells = [strip_accents(str(c)).lower().strip() for c in r]
        c_idx = next((idx for idx, v in enumerate(norm_cells) if any(k in v for k in (
            "ma khach hang", "ma nha cung cap", "ma doi tuong", "ma kh", "ma ncc", "ma dt"
        )) or v == "ma" or v == "code"), -1)
        n_idx = next((idx for idx, v in enumerate(norm_cells) if any(k in v for k in (
            "ten khach hang", "ten nha cung cap", "ten doi tuong", "ten kh", "ten ncc", "ten dt"
        )) or v == "ten" or v == "name"), -1)

        if c_idx != -1 and n_idx != -1:
            header_idx = i
            code_col, name_col = c_idx, n_idx
            mst_col = next((idx for idx, v in enumerate(norm_cells) if "ma so thue" in v or v == "mst"), -1)
            break

    if header_idx == -1:
        raise ValueError("Không tìm thấy cột 'Mã khách hàng/NCC' và 'Tên khách hàng/NCC' trong file!")

    items, seen_codes = [], set()
    for r in rows[header_idx + 1:]:
        if len(r) <= max(code_col, name_col):
            continue
        code = str(r[code_col]).strip()
        name = str(r[name_col]).strip()
        mst = str(r[mst_col]).strip() if (mst_col != -1 and mst_col < len(r)) else ""
        if not code or not name:
            continue
        if strip_accents(code).lower() in ("tong", "total", "cong"):
            continue
        key = (code.lower(), name.lower())
        if key in seen_codes:
            continue
        seen_codes.add(key)
        items.append({"code": code, "name": name, "mst": mst})

    return items


def extract_excel_or_csv(path):
    rows = _read_tabular_light(path)
    if not rows:
        raise ValueError("File Excel/CSV rong hoac khong doc duoc")

    header_idx = -1
    for i, r in enumerate(rows[:6]):
        lower = [str(c).replace("\ufeff", "").strip().lower() for c in r]
        if "trdt" in lower and ("acctccyamt" in lower or "trccyamt" in lower):
            header_idx = i
            break

    if header_idx != -1:
        col = {str(c).replace("\ufeff", "").strip().lower(): idx for idx, c in enumerate(rows[header_idx])}
        header_info, all_rows, stt = {}, [], 0
        beg_bal = end_bal = None
        for r in rows[header_idx + 1:]:
            if not any(r): continue
            d_val = normalize_date_ddmmyyyy(r[col['trdt']])
            if not re.match(r'^\d{2}/\d{2}/\d{4}$', d_val): continue
            stt += 1
            amt_col = 'acctccyamt' if 'acctccyamt' in col else 'trccyamt'
            amt = to_num(r[col[amt_col]])
            desc = str(r[col['rem']]).strip() if 'rem' in col and str(r[col['rem']]).strip() else ""
            if not desc and 'trcdnm' in col: desc = str(r[col['trcdnm']]).strip()
            bal = str(r[col['aftrbal']]).strip() if 'aftrbal' in col else ""
            if stt == 1 and 'bftrbal' in col: beg_bal = str(r[col['bftrbal']]).strip()
            end_bal = bal
            if not header_info.get('account_no') and 'idxactno' in col: header_info['account_no'] = str(r[col['idxactno']]).strip()
            if not header_info.get('customer') and 'name_4' in col: header_info['customer'] = str(r[col['name_4']]).strip()

            all_rows.append({
                'STT': str(stt), 'Ngay': d_val, 'MoTa': desc,
                'No': f"{abs(amt):,.0f}" if amt < 0 else "0", 'Co': f"{amt:,.0f}" if amt > 0 else "0",
                'SoDu': f"{to_num(bal):,.0f}" if bal else "",
                'DoiTac': "", 'MaGD': str(r[col.get('trcd', 0)]) if 'trcd' in col else "",
            })
        if all_rows:
            return {'bank': 'agribank', 'header_info': header_info, 'rows': all_rows,
                    'beginning_balance': f"{to_num(beg_bal):,.0f}" if beg_bal else None,
                    'ending_balance': f"{to_num(end_bal):,.0f}" if end_bal else None,
                    'total_debit': None, 'total_credit': None}

    vnd_rows = [r for r in rows if "VND" in [str(c).strip() for c in r]]
    if len(vnd_rows) >= 2:
        header_info, all_rows, stt = {}, [], 0
        for r in rows[:25]:
            line = " ".join(str(c) for c in r)
            m = re.search(r'(?:Số|Tài khoản):\s*([0-9A-Za-z]+)', line)
            if m and not header_info.get('account_no'): header_info['account_no'] = m.group(1)
        for r in vnd_rows:
            r_str = [str(c).strip() for c in r]
            v_idx = r_str.index("VND")
            desc = r_str[v_idx - 1] if v_idx > 0 else ""
            if "Loại tiền" in desc or "Loại tiền" in " ".join(r_str[:v_idx]): continue
            d_val = ""
            for v in r_str[:v_idx]:
                if re.search(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', v):
                    d_val = normalize_date_ddmmyyyy(v)
                    break
            if not d_val: continue
            no_str, co_str = "0", "0"
            for j in range(v_idx + 1, len(r_str)):
                val = r_str[j]
                if _is_money_token(val):
                    std_val = _vn_money_to_std_str(val)
                    if std_val != "0":
                        if j - v_idx <= 4: no_str = std_val
                        else: co_str = std_val
            if no_str != "0" or co_str != "0" or desc:
                stt += 1
                all_rows.append({'STT': str(stt), 'Ngay': d_val, 'MoTa': desc, 'No': no_str, 'Co': co_str, 'SoDu': "", 'DoiTac': "", 'MaGD': ""})
        if all_rows:
            return {'bank': 'sacombank', 'header_info': header_info, 'rows': all_rows, 'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None}

    all_rows, stt = [], 0
    for r in rows:
        r_str = [str(x).strip() for x in r]
        d_val = next((normalize_date_ddmmyyyy(x) for x in r_str if re.match(r'^(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})', x)), None)
        if not d_val: continue
        nums = [_vn_money_to_std_str(x) for x in r_str if _is_money_token(x) and _vn_money_to_std_str(x) != "0"]
        texts = [x for x in r_str if len(x) > 5 and not _is_money_token(x) and not re.match(r'^(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})', x)]
        if nums or texts:
            stt += 1
            all_rows.append({'STT': str(stt), 'Ngay': d_val, 'MoTa': texts[0] if texts else "", 'No': nums[0] if nums else "0", 'Co': nums[1] if len(nums) > 1 else "0", 'SoDu': "", 'DoiTac': "", 'MaGD': ""})
    return {'bank': 'excel_khac', 'header_info': {}, 'rows': all_rows, 'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None}


# =========================================================================
#  4. BO NHAN DIEN TAI KHOAN DOI UNG, CHU CONG TY, QUY LUAT HOC & SO KHOP MO
# =========================================================================
_STOP_WORDS_NAME = {
    "cong", "ty", "cty", "ct", "tnhh", "cp", "co", "phan", "mtv", "tm", "dv", "sx", "xd",
    "tmdv", "tmxd", "dvtm", "xnk", "va", "cn", "chi", "nhanh", "hcm", "hn", "ha", "nam",
    "ninh", "binh", "viet", "nam", "trach", "nhiem", "huu", "han", "thuong", "mai", "dich",
    "vu", "san", "xuat", "xay", "dung", "van", "tai", "dau", "tu", "phat", "trien", "tong",
    "hop", "tap", "doan", "hma", "nhh", "a/c", "ac", "tai", "tk", "so", "ct", "ck", "tt",
    "chuyen", "tien", "thanh", "toan", "tra", "hang", "cho", "den", "tu", "qua", "zalo",
    "mbvcb", "ibft", "napas", "qr", "gd", "ma", "hd", "hoa", "don", "mua", "ban", "nhom",
    "thep", "sat", "gach", "xi", "mang", "xang", "dau", "lop", "tien", "mat", "vang", "te",
    "don", "vi", "lpbank", "acb", "bidv", "vietinbank", "vcb", "mb", "agribank", "sacombank",
    "ho", "kinh", "doanh", "hkd", "ubnd", "xa", "phuong", "tinh", "huyen", "cap", "thoat", "nuoc"
}

_LEGAL_PREFIX_WORDS = {
    "cong", "ty", "cty", "ct", "tnhh", "cp", "co", "phan", "mtv", "ho", "kinh", "doanh", "hkd",
    "trach", "nhiem", "huu", "han", "tu", "van", "dau", "tu", "xay", "dung", "thuong", "mai",
    "dich", "vu", "san", "xuat", "van", "tai", "tm", "dv", "sx", "xd", "dt", "va", "dvtm",
    "tmdv", "tmxd", "xnk", "tap", "doan", "ubnd", "cap", "thoat", "nuoc", "dtxd", "tvxd", "xdtm", "sxkd"
}

_CANON_REPLACEMENTS = [
    (r'\bcong\s+ty\b', 'cty'),
    (r'\bct\s+(?=tnhh|cp|co\s+phan|tm|dv|dt|xd|sx)', 'cty '),
    (r'\bco\s+phan\b', 'cp'),
    (r'\btrach\s+nhiem\s+huu\s+han\b', 'tnhh'),
    (r'\bthuong\s+mai\b', 'tm'),
    (r'\bdich\s+vu\b', 'dv'),
    (r'\bxay\s+dung\b', 'xd'),
    (r'\bdau\s+tu\b', 'dt'),
    (r'\bsan\s+xuat\b', 'sx'),
    (r'\bvan\s+tai\b', 'vt'),
    (r'\bxuat\s+nhap\s+khau\b', 'xnk'),
    (r'\bchi\s+nhanh\b', 'cn'),
    (r'\bmtv\b', ''),
    (r'\bva\b', ''),
    (r'\btmdv\b', 'tm dv'),
    (r'\bdvtm\b', 'dv tm'),
    (r'\btmxd\b', 'tm xd'),
    (r'\bxdtm\b', 'xd tm'),
    (r'\bdtxd\b', 'dt xd'),
]

def _canon_biz(s):
    """Chuan hoa viet tat doanh nghiep hai chieu de khop chuan du ten hoi khac 1 ti"""
    norm = strip_accents(s or "").lower()
    norm = re.sub(r'[^a-z0-9\s]', ' ', norm)
    norm = re.sub(r'\s+', ' ', norm).strip()
    for pat, rep in _CANON_REPLACEMENTS:
        norm = re.sub(pat, rep, norm)
    # Cat bo phan chi nhanh o cuoi neu co ("cn tinh ninh binh", "cn ha nam", "hb")
    norm = re.sub(r'\bcn\s+(?:tinh\s+)?(?:ninh binh|ha nam|hung yen|phu tho|ha noi|viet nam).*$', '', norm).strip()
    norm = re.sub(r'\s+', ' ', norm).strip()
    return norm

def _owner_code_candidates(owner_name, owner_code=""):
    cands = set()
    if owner_code:
        cands.add(owner_code.lower())
    if not owner_name:
        return cands
    clean = strip_accents(owner_name).lower()
    clean = re.sub(r'[^a-z0-9\s]', ' ', clean)
    tokens = [w for w in clean.split() if w and w not in _STOP_WORDS_NAME and len(w) >= 2]
    if len(tokens) >= 2:
        cands.add(" ".join(tokens))
        cands.add("".join(tokens))
        cands.add(" ".join(tokens[-2:]))
        cands.add("".join(tokens[-2:]))
    return cands

def _normalize_partner_search_variants(name):
    raw_norm = re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9\s]', ' ', strip_accents(name).lower())).strip()
    canon = _canon_biz(name)

    variants = []
    seen = set()
    def add_v(v, score):
        v = v.strip()
        if v and v not in seen:
            seen.add(v)
            variants.append((v, score))

    if len(raw_norm.split()) >= 2:
        add_v(raw_norm, len(raw_norm) + 35)
    if len(canon.split()) >= 2:
        add_v(canon, len(canon) + 30)

    # Rut gon tien to phap ly
    for base in (canon, raw_norm):
        words = base.split()
        core_words = []
        for idx, w in enumerate(words):
            if w in _LEGAL_PREFIX_WORDS and (len(words) - idx) > 2:
                continue
            core_words.append(w)
        core = " ".join(core_words).strip()
        if core and len(core) >= 4:
            add_v(core, len(core) + 15)
        if len(core_words) >= 3:
            tail2 = " ".join(core_words[-2:])
            if len(tail2) >= 5:
                add_v(tail2, len(tail2) + 2)
            tail3 = " ".join(core_words[-3:])
            if len(tail3) >= 7:
                add_v(tail3, len(tail3) + 5)

    # Core tokens dung cho so khop mo (token overlap)
    dist_tokens = [w for w in canon.split() if w not in _LEGAL_PREFIX_WORDS and w not in _STOP_WORDS_NAME and len(w) >= 2]
    return variants, dist_tokens

def _match_from_directory(search_text_norm, canon_text, raw_text, partner_list, owner_cands):
    if not partner_list:
        return None

    best_code = None
    best_score = -1

    masked_raw = " " + search_text_norm + " "
    masked_canon = " " + canon_text + " "
    for oc in sorted(owner_cands, key=len, reverse=True):
        if len(oc) >= 4:
            masked_raw = masked_raw.replace(f" {oc} ", " ")
            masked_canon = masked_canon.replace(f" {oc} ", " ")

    masked_tokens = set(masked_canon.split())
    fuzzy_candidates = []

    for item in partner_list:
        code = str(item.get("code", "")).strip()
        name = str(item.get("name", "")).strip()
        mst = re.sub(r'\s+', '', str(item.get("mst", "") or ""))
        if not code or not name:
            continue

        # 1. Khop Ma so thue
        if len(mst) >= 10 and mst in raw_text:
            return code

        # 2. Khop chinh xac cum ten (ca ban goc va ban chuan hoa viet tat)
        variants, dist_tokens = _normalize_partner_search_variants(name)
        matched_exact = False
        for variant, base_score in variants:
            if variant in owner_cands:
                continue
            pat = r'(?<![a-z0-9])' + re.escape(variant) + r'(?![a-z0-9])'
            if re.search(pat, masked_canon) or re.search(pat, masked_raw):
                matched_exact = True
                if base_score > best_score:
                    best_score = base_score
                    best_code = code

        # 3. Thu so khop mo theo tu phan biet (de bat truong hop ten khac 1 ti)
        if not matched_exact and len(dist_tokens) >= 2:
            hits = [t for t in dist_tokens if t in masked_tokens]
            ratio = len(hits) / len(dist_tokens)
            # Phai trung it nhat 2 tu phan biet VA >= 66% tong so tu phan biet
            if len(hits) >= 2 and ratio >= 0.66:
                fuzzy_score = len(hits) * 10 + ratio * 10
                fuzzy_candidates.append((fuzzy_score, code))

    if best_code:
        return best_code

    if fuzzy_candidates:
        fuzzy_candidates.sort(key=lambda x: x[0], reverse=True)
        return fuzzy_candidates[0][1]

    return None

def _resolve_bank_partner_code(bank_key, source_file=""):
    known = {"acb", "agribank", "bidv", "vietinbank", "lpbank", "mbbank", "vpbank", "sacombank"}
    bk = str(bank_key or "").strip().lower()
    if bk in known:
        return bk
    fn = strip_accents(source_file).lower()
    for b in known:
        if b in fn:
            return b
    return ""


def classify_for_asp(row, owner_name="", owner_code="", owner_acc="",
                     customers=None, suppliers=None, owners=None, rules=None,
                     bank_code="", acc_to_comp=None):
    no_val = to_num(row.get('No'))
    co_val = to_num(row.get('Co'))
    mota = row.get('MoTa', '')
    doitac = row.get('DoiTac', '')
    magd = str(row.get('MaGD', '') or '').strip().upper()

    txt = strip_accents(mota).lower()
    dt_txt = strip_accents(doitac).lower()
    combined = re.sub(r'\s+', ' ', f"{txt} {dt_txt}").strip()
    canon_combined = _canon_biz(combined)

    # -1. UU TIEN CAO NHAT: QUY LUAT DA LUU CUA RIENG CONG TY (Sua 1 lan, lan sau tu theo!)
    if rules:
        direction = "in" if (co_val > 0 and no_val == 0) else ("out" if (no_val > 0 and co_val == 0) else "both")
        for rule in reversed(rules):
            kw = strip_accents(rule.get("keyword", "")).lower().strip()
            r_dir = rule.get("direction", "both")
            if not kw:
                continue
            if r_dir != "both" and r_dir != direction:
                continue
            if kw in combined or _canon_biz(kw) in canon_combined:
                r_tk = str(rule.get("tk", "")).strip()
                r_madt = str(rule.get("madt", "")).strip()
                if r_tk in ("6422", "515") and not r_madt:
                    r_madt = bank_code
                return r_tk, r_madt

    # 0. CAC KHOAN VAY / THUE / BHXH / NOI BO / LUONG -> DE TRONG ""
    other_non_keywords = (
        "thu goc", "thu lai", "thu lai phat", "khe uoc", "tra von vay", "tra lai vay",
        "loan repayment", "thau chi", "dong goc lai", "giai ngan",
        "nop thue", "thue thu nhap", "thue dien tu", "cham nop thue", "kho bac", "kbnn", "kb:1327",
        "bhxh", "bao hiem", "chuyen noi bo", "chi luong",
        "ve duong", "vetc", "chuyen tra lai tien", "nop nham", "giao dich nham"
    )
    if any(k in combined for k in other_non_keywords):
        return "", ""

    # 1. LAI TIEN GUI (TK 515) -> MaDT = ten ngan hang
    if co_val > 0 and no_val == 0:
        if (magd == "IIPD" or any(k in combined for k in (
            "lai nhap von", "but toan tra lai", "tra lai tai khoan", "tra lai tien gui",
            "chi lai tk", "lai tien gui", "thanh toan lai thang"
        )) or re.match(r'^tra lai\b', txt.strip())):
            return "515", bank_code

    # 2. PHI NGAN HANG / PHI CHUYEN TIEN (TK 6422) -> MaDT = ten ngan hang
    if no_val > 0 and co_val == 0:
        is_fee = False
        if magd == "DDM" or magd.startswith("CHG"):
            is_fee = True
        elif any(k in combined for k in (
            "thu phi", "phi mb", "phi dv", "phi bdsd", "phi sms", "phi bsms",
            "phi qltk", "phi quan ly", "phi dich vu", "phi duy tri", "phi phat hanh",
            "phi ban sec", "phi chuyen tien", "phi ck", "phi hoi bao", "service charge", "services charge"
        )):
            is_fee = True
        elif txt.strip() == "vat" or ("vat" in txt and "phi dich vu" in dt_txt):
            is_fee = True
        elif "outward remittance" in magd.lower() and no_val <= 300000:
            is_fee = True

        if is_fee:
            return "6422", bank_code

    # 3. CHU CONG TY / NGUOI NOP-RUT QUY & NOP/RUT TIEN MAT, CHI SEC (TK 1111) -> MaDT de trong
    owner_people = [strip_accents(p).lower().strip() for p in (owners or []) if p and str(p).strip()]
    is_owner_person = any(
        re.search(r'(?<![a-z0-9])' + re.escape(pname) + r'(?![a-z0-9])', combined)
        for pname in owner_people if len(pname) >= 3
    )
    is_cash_or_cheque = (
        "tien mat, vang tien te tai don vi" in dt_txt
        or is_owner_person
        or any(k in combined for k in (
            "nop tien", "rut tien", "rut sec", "chi sec", "linh tien",
            "nt vao tk", "nop tm", "nt-tk-acb", "nop tien mat"
        ))
        or txt.strip() in ("nt", "rt")
        or (bool(re.search(r'\b(rt|nt)\b', txt)) and not any(k in txt for k in ("cty", "cong ty", "tnhh", "hang")))
    )
    if is_cash_or_cheque:
        return "1111", ""

    owner_cands = _owner_code_candidates(owner_name, owner_code=owner_code)
    raw_combined = f"{mota} {doitac}"
    norm_combined = re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9\s]', ' ', strip_accents(raw_combined).lower())).strip()

    # 4. KHACH TRA TIEN (TK 131) -> CHI TIM TRONG DANH SACH KHACH HANG CUA DUNG CONG TY DO
    if co_val > 0 and no_val == 0:
        ma_dt = _match_from_directory(norm_combined, canon_combined, raw_combined, customers or [], owner_cands) or ""
        if not ma_dt and acc_to_comp:
            for acc_no, comp_code in acc_to_comp.items():
                if len(acc_no) >= 7 and acc_no != owner_acc and comp_code not in owner_cands and acc_no in raw_combined:
                    ma_dt = comp_code
                    break
        is_customer_pay = (
            bool(ma_dt)
            or bool(dt_txt and "tien mat" not in dt_txt)
            or any(k in txt for k in (
                "cty", "cong ty", "tnhh", "cp ", "co phan", "thanh toan", "tra tien",
                "chuyen tien", "chuyen khoan", "ck ", "tien hang", "cong no", "hoa don",
                "hd ", "dat hang", "dat coc", "ung tien", "hop dong", "van chuyen",
                "tt ", "ttcn", "ban hang", "mbvcb", "247 den", "ibft", "qr"
            ))
        )
        if is_customer_pay:
            return "131", ma_dt

    # 5. TRA TIEN NHA CUNG CAP (TK 331) -> CHI TIM TRONG DANH SACH NHA CUNG CAP CUA DUNG CONG TY DO
    if no_val > 0 and co_val == 0:
        ma_dt = _match_from_directory(norm_combined, canon_combined, raw_combined, suppliers or [], owner_cands) or ""
        if not ma_dt and acc_to_comp:
            for acc_no, comp_code in acc_to_comp.items():
                if len(acc_no) >= 7 and acc_no != owner_acc and comp_code not in owner_cands and acc_no in raw_combined:
                    ma_dt = comp_code
                    break
        is_supplier_pay = (
            bool(ma_dt)
            or bool(dt_txt and "tien mat" not in dt_txt)
            or any(k in txt for k in (
                "cty", "cong ty", "tnhh", "cp ", "co phan", "thanh toan", "tra tien",
                "chuyen tien", "chuyen khoan", "ck ", "tien hang", "mua hang", "tien thep",
                "tien sat", "tien nhom", "tien gach", "tien xi mang", "tien lop", "tien dau",
                "phu tung", "dat coc", "tt ", "247 di", "napas", "hoa don", "hd"
            ))
        )
        if is_supplier_pay:
            return "331", ma_dt

    return "", ""


# =========================================================================
#  5. DIEU PHOI, DOI CHIEU & XUAT EXCEL
# =========================================================================
def extract_any(path, password=None):
    ext = os.path.splitext(path)[1].lower()
    source_name = os.path.basename(path)

    if ext in (".html", ".htm", ".xls"):
        try:
            with open(path, "rb") as f:
                head = f.read(600).decode("utf-8", errors="ignore").strip().lower()
            if any(tag in head for tag in ("<html", "<table", "<body", "<!doctype")):
                return extract_html(path)
        except Exception:
            pass

    if ext in (".html", ".htm"): return extract_html(path)
    if ext in (".xlsx", ".xls", ".csv"):
        try:
            res = extract_excel_or_csv(path)
            res['source_file'], res['loi'] = source_name, None
            return res
        except Exception as e:
            return {'source_file': source_name, 'bank': 'LOI', 'header_info': {}, 'rows': [], 'loi': f"Loi Excel/CSV: {e}"}

    try:
        pdf = pdfplumber.open(path, password=password or "")
    except Exception as e:
        return {'source_file': source_name, 'bank': 'LOI', 'header_info': {}, 'rows': [], 'loi': f"Khong mo duoc PDF: {e}"}

    with pdf:
        txt = (pdf.pages[0].extract_text() or "") + "\n" + (pdf.pages[-1].extract_text() or "" if len(pdf.pages) > 1 else "")
        bank = detect_bank(txt)
        try:
            if bank == "vietinbank": res = extract_vietinbank(pdf)
            elif bank == "bidv": res = extract_bidv(pdf)
            elif bank == "lpbank": res = extract_lpbank(pdf)
            elif bank == "acb": res = extract_acb(pdf)
            elif bank == "mbbank": res = extract_mbbank(pdf)
            else: return {'source_file': source_name, 'bank': 'khong_xac_dinh', 'header_info': {}, 'rows': [], 'loi': 'Chua ho tro PDF ngan hang nay'}
        except Exception as e:
            return {'source_file': source_name, 'bank': bank, 'header_info': {}, 'rows': [], 'loi': f"Loi trich xuat PDF: {e}"}

    res['source_file'], res['bank'], res['loi'] = source_name, bank, None
    return res


def validate(res):
    warnings = []
    rows = res['rows']
    calc_debit = sum(to_num(r['No']) for r in rows)
    calc_credit = sum(to_num(r['Co']) for r in rows)
    if res.get('total_debit') and abs(calc_debit - to_num(res['total_debit'])) > 0.01:
        warnings.append(f"Tong No khong khop: tinh {calc_debit:,.0f} vs sao ke {res['total_debit']}")
    if res.get('total_credit') and abs(calc_credit - to_num(res['total_credit'])) > 0.01:
        warnings.append(f"Tong Co khong khop: tinh {calc_credit:,.0f} vs sao ke {res['total_credit']}")
    if res.get('beginning_balance') and any(r.get('SoDu') for r in rows):
        bal = to_num(res['beginning_balance'])
        for r in rows:
            bal = bal - to_num(r['No']) + to_num(r['Co'])
            if r.get('SoDu') and abs(bal - to_num(r['SoDu'])) > 0.01:
                warnings.append(f"Dong STT {r['STT']}: so du tinh {bal:,.0f} != sao ke {to_num(r['SoDu']):,.0f}")
    if not rows and not res.get('loi'):
        warnings.append("Khong trich xuat duoc dong giao dich nao")
    return warnings


def make_sheet_name(file_name, used):
    stem = re.sub(r'[\[\]:*?/\\]', ' ', os.path.splitext(file_name)[0]).strip().strip("'")
    stem = re.sub(r'\s+', ' ', stem) or "Sheet"
    base = stem[:31]
    name, n = base, 1
    while name.lower() in used:
        n += 1
        suffix = f" ({n})"
        name = base[:31 - len(suffix)] + suffix
    used.add(name.lower())
    return name


def create_workbook_from_sheets(sheets_data, mode="4col"):
    FONT = "Arial"
    header_font = Font(name=FONT, bold=True, color="2C3E50", size=11)
    header_fill = PatternFill("solid", fgColor="DCE6F7" if mode == "4col" else "E6DEF7")
    normal_font = Font(name=FONT, size=10)
    thin = Side(style="thin", color="D0D7E2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_fmt = '#,##0'

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()

    if mode == "asp":
        headers = ["Ngày", "Nợ", "Có", "Diễn giải", "Tài khoản đối ứng", "Mã đối tượng"]
        widths = [15, 17, 17, 60, 20, 22]
    else:
        headers = ["Ngày", "Nợ", "Có", "Diễn giải"]
        widths = [15, 18, 18, 65]

    for s in sheets_data:
        name = make_sheet_name(s['file'], used_names)
        ws = wb.create_sheet(name)

        for c, h in enumerate(headers, start=1):
            cell = ws.cell(1, c, h)
            cell.font, cell.fill, cell.border = header_font, header_fill, border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        if s.get('error'):
            ws.cell(2, 1, "LỖI")
            ws.cell(2, 4, s['error'])
        else:
            for r_idx, gd in enumerate(s['rows'], start=2):
                no_num = to_num(gd.get('No'))
                co_num = to_num(gd.get('Co'))

                c_ngay = ws.cell(r_idx, 1, gd.get('Ngay', ''))
                c_no = ws.cell(r_idx, 2, no_num if no_num != 0 else None)
                c_co = ws.cell(r_idx, 3, co_num if co_num != 0 else None)
                c_dg = ws.cell(r_idx, 4, gd.get('MoTa', ''))

                c_ngay.alignment = Alignment(horizontal="center", vertical="top")
                c_no.alignment = Alignment(horizontal="right", vertical="top")
                c_co.alignment = Alignment(horizontal="right", vertical="top")
                c_dg.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

                cells = [c_ngay, c_no, c_co, c_dg]

                if mode == "asp":
                    c_tk = ws.cell(r_idx, 5, gd.get('TK', ''))
                    c_dt = ws.cell(r_idx, 6, gd.get('MaDT', ''))
                    c_tk.alignment = Alignment(horizontal="center", vertical="top")
                    c_dt.alignment = Alignment(horizontal="left", vertical="top")
                    cells.extend([c_tk, c_dt])

                for cell in cells:
                    cell.font = normal_font
                    cell.border = border
                if c_no.value is not None: c_no.number_format = num_fmt
                if c_co.value is not None: c_co.number_format = num_fmt

        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    if not wb.sheetnames:
        ws = wb.create_sheet("SaoKe")
        for c, h in enumerate(headers, start=1):
            ws.cell(1, c, h)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =========================================================================
#  6. GIAO DIEN WEB PASTEL + BO LOC CHU CONG TY + HOC QUY LUAT TU DONG
# =========================================================================
HTML_UI = r"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Công cụ trích xuất sao kê ngân hàng</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, sans-serif;
            margin: 0;
            padding: 24px;
            background: linear-gradient(135deg, #f3f6ff 0%, #fdf2f8 50%, #eefaf4 100%);
            min-height: 100vh;
            color: #3d485c;
        }
        .container {
            max-width: 1280px;
            margin: 0 auto;
            background: rgba(255, 255, 255, 0.96);
            padding: 32px;
            border-radius: 20px;
            box-shadow: 0 12px 35px rgba(130, 145, 185, 0.16);
            border: 1px solid #eaeef7;
        }
        h1 { color: #4b659e; margin: 0 0 8px 0; font-size: 25px; letter-spacing: -0.3px; }
        .subtitle { color: #6b768c; font-size: 14px; margin-bottom: 20px; line-height: 1.6; }

        .partner-card {
            background: linear-gradient(180deg, #faf7ff 0%, #f4f7ff 100%);
            border: 1.5px solid #e1dbf7;
            border-radius: 16px;
            padding: 18px 22px;
            margin-bottom: 20px;
        }
        .partner-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 12px;
        }
        .partner-row { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
        select {
            padding: 9px 12px;
            border: 1.5px solid #d5ddf0;
            border-radius: 10px;
            font-size: 13.5px;
            background: #fff;
            color: #3d485c;
            min-width: 310px;
            max-width: 100%;
        }
        .chip-list { display: flex; gap: 7px; flex-wrap: wrap; align-items: center; margin: 8px 0; }
        .acc-chip {
            display: inline-flex; align-items: center; gap: 6px;
            background: #eef3ff; border: 1px solid #cfe0fc; color: #3c5791;
            padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 600;
        }
        .owner-chip {
            display: inline-flex; align-items: center; gap: 6px;
            background: #fff3e0; border: 1px solid #f7d5a8; color: #8c5214;
            padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 600;
        }
        .chip-del { cursor: pointer; color: #c94c5b; font-weight: bold; border: none; background: transparent; font-size: 14px; padding: 0 2px; }
        .badge-count {
            display: inline-block; padding: 5px 11px; border-radius: 8px; font-size: 12.5px; font-weight: 600;
            background: #e8f8f0; color: #1f704b; border: 1px solid #bfead3;
        }
        .badge-empty { background: #f5f6fa; color: #7c879c; border-color: #e0e4ee; }
        .partner-group {
            display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
            background: #fff; padding: 8px 12px; border-radius: 12px; border: 1px solid #e6ebf5;
        }

        .upload-zone {
            border: 2.5px dashed #9cb8f3; background: #f4f8ff; padding: 26px 20px;
            text-align: center; border-radius: 16px; cursor: pointer; margin-bottom: 18px; transition: all 0.2s ease;
        }
        .upload-zone:hover { background: #e9f1ff; border-color: #779be8; transform: translateY(-1px); }
        .controls { display: flex; gap: 14px; align-items: center; flex-wrap: wrap; margin-bottom: 16px; }
        input[type="password"], input[type="text"] {
            padding: 10px 14px; border: 1.5px solid #d9e2f2; border-radius: 10px;
            font-size: 13.5px; background: #fbfcff; color: #3d485c; outline: none;
        }
        input:focus { border-color: #8eb0f2; }

        .btn {
            padding: 10px 20px; border: none; border-radius: 12px; font-weight: 700; font-size: 13.5px;
            cursor: pointer; text-decoration: none; display: inline-block; transition: all 0.15s ease; transform: translateY(0);
        }
        .btn-sm { padding: 7px 12px; font-size: 12px; border-radius: 9px; }
        .btn-xs { padding: 4px 8px; font-size: 11.5px; border-radius: 7px; }
        .btn-primary {
            background: linear-gradient(180deg, #91b3fa 0%, #769ce9 100%); color: #fff;
            box-shadow: 0 5px 0 #5b7fc9, 0 8px 16px rgba(118, 156, 233, 0.32);
        }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 7px 0 #5b7fc9, 0 12px 20px rgba(118, 156, 233, 0.38); }
        .btn-primary:active { transform: translateY(3px); box-shadow: 0 2px 0 #5b7fc9, 0 4px 8px rgba(118, 156, 233, 0.25); }

        .btn-download {
            background: linear-gradient(180deg, #83ddb8 0%, #63c79e 100%); color: #164f38;
            box-shadow: 0 5px 0 #48a67f, 0 8px 16px rgba(99, 199, 158, 0.32);
        }
        .btn-download:hover { transform: translateY(-2px); box-shadow: 0 7px 0 #48a67f, 0 12px 20px rgba(99, 199, 158, 0.40); }
        .btn-download:active { transform: translateY(3px); box-shadow: 0 2px 0 #48a67f, 0 4px 8px rgba(99, 199, 158, 0.25); }

        .btn-asp {
            background: linear-gradient(180deg, #c2aff7 0%, #a58ee8 100%); color: #fff;
            box-shadow: 0 5px 0 #8269c9, 0 8px 16px rgba(165, 142, 232, 0.34);
        }
        .btn-asp:hover { transform: translateY(-2px); box-shadow: 0 7px 0 #8269c9, 0 12px 20px rgba(165, 142, 232, 0.42); }
        .btn-asp:active { transform: translateY(3px); box-shadow: 0 2px 0 #8269c9, 0 4px 8px rgba(165, 142, 232, 0.25); }

        .btn-soft {
            background: linear-gradient(180deg, #ffffff 0%, #f1f4fb 100%); color: #4b5d7a;
            border: 1px solid #d9e1f2; box-shadow: 0 3px 0 #d0d8ea;
        }
        .btn-soft:hover { transform: translateY(-1px); box-shadow: 0 4px 0 #c3cde3; }
        .btn-soft:active { transform: translateY(2px); box-shadow: 0 1px 0 #c3cde3; }

        .btn-mint {
            background: linear-gradient(180deg, #b8f0d8 0%, #90dfb9 100%); color: #164f38;
            box-shadow: 0 3px 0 #68bf96;
        }
        .btn-mint:hover { transform: translateY(-1px); box-shadow: 0 4px 0 #68bf96; }
        .btn-mint:active { transform: translateY(2px); box-shadow: 0 1px 0 #68bf96; }

        .btn-peach {
            background: linear-gradient(180deg, #ffd4be 0%, #f8b799 100%); color: #693018;
            box-shadow: 0 3px 0 #dc9372;
        }
        .btn-peach:hover { transform: translateY(-1px); box-shadow: 0 4px 0 #dc9372; }
        .btn-peach:active { transform: translateY(2px); box-shadow: 0 1px 0 #dc9372; }

        .btn-amber {
            background: linear-gradient(180deg, #ffe7b8 0%, #f9cf7e 100%); color: #664407;
            box-shadow: 0 3px 0 #d9aa52;
        }
        .btn-amber:hover { transform: translateY(-1px); box-shadow: 0 4px 0 #d9aa52; }
        .btn-amber:active { transform: translateY(2px); box-shadow: 0 1px 0 #d9aa52; }

        .tab-btn {
            padding: 9px 16px; border: 1px solid #dce3f0; background: #f5f7fc; color: #4b5d7a;
            border-radius: 10px; cursor: pointer; font-size: 13px; font-weight: 600;
            box-shadow: 0 3px 0 #d5dceb; transition: all 0.15s ease;
        }
        .tab-btn:hover { transform: translateY(-1px); background: #ebf0fb; box-shadow: 0 4px 0 #c9d2e6; }
        .tab-btn:active { transform: translateY(2px); box-shadow: 0 1px 0 #c9d2e6; }
        .tab-btn.active {
            background: linear-gradient(180deg, #b8a9f5 0%, #9c8be8 100%); color: #fff;
            border-color: #8b78db; box-shadow: 0 3px 0 #7562c7, 0 5px 10px rgba(156, 139, 232, 0.28);
        }

        .mode-switch { display: inline-flex; gap: 8px; background: #f2f5fc; padding: 6px; border-radius: 12px; border: 1px solid #e1e8f5; }
        .mode-pill { padding: 7px 14px; border-radius: 8px; border: none; background: transparent; color: #5b6982; font-weight: 600; font-size: 13px; cursor: pointer; }
        .mode-pill.active { background: #fff; color: #4b659e; box-shadow: 0 2px 6px rgba(91, 127, 201, 0.18); }

        #fileList { margin: 10px 0 0 0; font-size: 13px; color: #5b7fc9; font-weight: 600; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; background: #fff; }
        th, td { border: 1px solid #e6ebf5; padding: 8px 10px; text-align: left; }
        th { background: #e4ecfc; color: #364b75; font-weight: 700; position: sticky; top: 0; z-index: 2; }
        tbody tr:nth-child(even) { background: #fafbfe; }
        tbody tr:hover { background: #f2f6ff; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .debit { color: #d45b68; font-weight: 600; }
        .credit { color: #39996d; font-weight: 600; }
        .ok { background: #e6f9ef; color: #227850; font-weight: 700; padding: 4px 10px; border-radius: 6px; display: inline-block; }
        .warn { background: #fde9eb; color: #b8424d; font-weight: 700; padding: 4px 10px; border-radius: 6px; display: inline-block; }
        .loader { display: none; color: #6787cc; font-weight: 700; margin: 14px 0; }
        .preview-section { margin-top: 26px; border-top: 2px dashed #e8edf7; padding-top: 20px; }
        .tabs { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 14px; }
        .preview-toolbar { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 12px; }
        .scroll-box { max-height: 500px; overflow: auto; border: 1px solid #e1e7f2; border-radius: 12px; }
        .cell-input { width: 100%; padding: 5px 7px; border: 1px solid #dde4f2; border-radius: 6px; font-size: 12.5px; background: #fcfdff; color: #3d485c; }
        .cell-input:focus { border-color: #9c8be8; background: #fff; outline: none; }
        .tk-badge { text-align: center; font-weight: 700; color: #5b48a8; }

        .modal-backdrop {
            display: none; position: fixed; inset: 0; background: rgba(50, 62, 90, 0.4);
            backdrop-filter: blur(3px); z-index: 999; align-items: center; justify-content: center; padding: 16px;
        }
        .modal-box {
            background: #fff; border-radius: 18px; padding: 24px; width: 100%; max-width: 540px;
            box-shadow: 0 20px 50px rgba(60, 75, 110, 0.25); border: 1px solid #e5ebf7;
        }
        .modal-box.wide { max-width: 860px; }
        .modal-title { font-size: 18px; font-weight: 700; color: #4b659e; margin: 0 0 14px 0; }
        .form-group { margin-bottom: 12px; }
        .form-group label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 5px; color: #4b5d7a; }
        .form-group input, .form-group select { width: 100%; padding: 9px 12px; border: 1.5px solid #d9e2f2; border-radius: 10px; font-size: 13.5px; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px; }
    </style>
</head>
<body>
<div class="container">
    <h1>🌸 Công cụ trích xuất sao kê ngân hàng</h1>
    <div class="subtitle">
        Hỗ trợ: <b>VietinBank, BIDV, LPBank, ACB, MBBank (PDF)</b> • <b>VPBank (HTML)</b> • <b>Sacombank, Agribank (Excel/CSV)</b>.<br>
        Chuẩn ngày: <b>dd/mm/yyyy</b> • Mỗi công ty có riêng Danh bạ KH/NCC, Chủ công ty (1111) và Quy luật đã học (sửa 1 lần, lần sau tự theo).
    </div>

    <!-- KHU VUC QUAN LY CONG TY, SO TK, CHU CONG TY & DANH SACH KH/NCC -->
    <div class="partner-card">
        <div class="partner-header">
            <div style="font-weight: 700; color: #574699; font-size: 15px;">
                🏢 Quản lý Công ty, Số TK, Chủ công ty (1111) &amp; Danh bạ KH / NCC
            </div>
            <div id="syncBadge" style="font-size: 12px; color: #1f704b; font-weight: 600;"></div>
        </div>

        <div class="partner-row" style="margin-bottom: 8px;">
            <label style="font-weight: 600; font-size: 13px;">Công ty:</label>
            <select id="companySelect" onchange="onSelectCompany()"></select>
            <button class="btn btn-primary btn-sm" onclick="openAddCompanyModal()">➕ Thêm mới Công ty</button>
            <button class="btn btn-soft btn-sm" onclick="openAddAccountModal()">🏦 Thêm TK Ngân hàng</button>
            <button class="btn btn-amber btn-sm" onclick="openViewRulesModal()" id="rulesBtn">📜 Quy luật đã học (0)</button>
            <button class="btn btn-peach btn-sm" onclick="deleteCompany()">🗑️ Xóa CT</button>
        </div>

        <div id="accChipsBox" class="chip-list"></div>
        <div id="ownerChipsBox" class="chip-list"></div>

        <div class="partner-row" style="margin-top: 10px;">
            <div class="partner-group">
                <span id="khStatus" class="badge-count badge-empty">0 Khách hàng</span>
                <input type="file" id="upKhReplace" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadPartnerList('customers', 'replace', this)">
                <input type="file" id="upKhUpdate" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadPartnerList('customers', 'update', this)">
                <button class="btn btn-soft btn-sm" onclick="triggerPartnerUpload('upKhReplace')">📥 Nạp mới DS KH</button>
                <button class="btn btn-mint btn-sm" onclick="triggerPartnerUpload('upKhUpdate')">🔄 Update DS KH</button>
                <button class="btn btn-soft btn-sm" onclick="openViewPartnersModal('customers')">👁️ Xem</button>
            </div>

            <div class="partner-group">
                <span id="nccStatus" class="badge-count badge-empty">0 Nhà cung cấp</span>
                <input type="file" id="upNccReplace" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadPartnerList('suppliers', 'replace', this)">
                <input type="file" id="upNccUpdate" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadPartnerList('suppliers', 'update', this)">
                <button class="btn btn-soft btn-sm" onclick="triggerPartnerUpload('upNccReplace')">📥 Nạp mới DS NCC</button>
                <button class="btn btn-mint btn-sm" onclick="triggerPartnerUpload('upNccUpdate')">🔄 Update DS NCC</button>
                <button class="btn btn-soft btn-sm" onclick="openViewPartnersModal('suppliers')">👁️ Xem</button>
            </div>
        </div>
    </div>

    <div class="upload-zone" onclick="document.getElementById('files').click()">
        <input type="file" id="files" multiple accept=".pdf,.html,.htm,.xlsx,.xls,.csv" style="display:none" onchange="showFiles()">
        <div style="font-size: 16px; font-weight: bold; color: #5b7fc9;">📂 Bấm vào đây để chọn 1 hoặc nhiều file sao kê</div>
        <div id="fileList"></div>
    </div>

    <div class="controls">
        <input type="password" id="pwd" style="width: 320px; max-width: 100%;" placeholder="🔑 Mật khẩu file PDF (nếu có)">
        <button class="btn btn-primary" onclick="uploadAndProcess()">✨ BẮT ĐẦU TRÍCH XUẤT &amp; XEM TRƯỚC</button>
    </div>

    <div id="loader" class="loader">⏳ Đang trích xuất, đối chiếu số dư và ghép mã đối tượng theo đúng công ty...</div>

    <div id="resultBox" style="display:none;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 14px;">
            <h3 style="margin: 0; color: #4b659e;">📊 Bảng Đối Chiếu Tổng Quan</h3>
            <div style="display: flex; gap: 12px; flex-wrap: wrap;">
                <button class="btn btn-download" onclick="downloadExcel('4col')">📥 Tải Excel (4 cột)</button>
                <button class="btn btn-asp" onclick="downloadExcel('asp')">📘 Kết xuất cho ASP (.xlsx)</button>
            </div>
        </div>

        <div style="border-radius: 12px; overflow: hidden; border: 1px solid #e6ebf5;">
            <table>
                <thead>
                    <tr>
                        <th>File nguồn</th><th>Ngân hàng</th><th>Số TK</th>
                        <th class="text-right">Số GD</th><th class="text-right">Tổng Nợ</th>
                        <th class="text-right">Tổng Có</th><th>Đối chiếu</th>
                    </tr>
                </thead>
                <tbody id="summaryBody"></tbody>
            </table>
        </div>

        <div class="preview-section">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 12px;">
                <h3 style="margin: 0; color: #4b659e;">🔍 Xem Trước &amp; Sửa Quy Luật (Sửa 1 lần, lần sau tự theo)</h3>
                <div class="mode-switch">
                    <button id="btnMode4" class="mode-pill" onclick="switchPreviewMode('4col')">Mẫu 4 cột</button>
                    <button id="btnModeAsp" class="mode-pill active" onclick="switchPreviewMode('asp')">Kết xuất cho ASP (6 cột)</button>
                </div>
            </div>

            <div id="tabs" class="tabs"></div>

            <div class="preview-toolbar">
                <div id="sheetMeta" style="font-size: 13px; color: #5b6982; font-weight: 500;"></div>
                <input type="text" id="searchBox" style="width: 300px;" placeholder="🔎 Tìm ngày, số tiền, diễn giải, TK, mã..." oninput="filterPreview()">
            </div>

            <datalist id="partnerDatalist"></datalist>
            <div class="scroll-box">
                <table>
                    <thead id="previewHead"></thead>
                    <tbody id="previewBody"></tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<!-- MODAL 1: THEM MOI CONG TY -->
<div id="modalAddCompany" class="modal-backdrop">
    <div class="modal-box">
        <div class="modal-title">➕ Thêm mới Công ty &amp; Tài khoản Ngân hàng</div>
        <div class="form-group"><label>Tên công ty (*):</label><input type="text" id="newCompName" placeholder="VD: CÔNG TY TNHH ABC"></div>
        <div class="form-group"><label>Mã công ty - viết liền không dấu (*):</label><input type="text" id="newCompCode" placeholder="VD: congtyabc"></div>
        <div class="form-group">
            <label>Ngân hàng:</label>
            <select id="newCompBank">
                <option value="Vietinbank">Vietinbank</option><option value="BIDV">BIDV</option><option value="LPBank">LPBank</option>
                <option value="ACB">ACB</option><option value="MB">MBBank</option><option value="Agribank">Agribank</option>
                <option value="Sacombank">Sacombank</option><option value="Vietcombank">Vietcombank</option><option value="VPBank">VPBank</option>
                <option value="MSB">MSB</option><option value="NCB">NCB</option><option value="Co-op Bank">Co-op Bank</option><option value="Khác">Khác</option>
            </select>
        </div>
        <div class="form-group"><label>Số tài khoản ngân hàng (có thể để trống):</label><input type="text" id="newCompAcc" placeholder="VD: 118003052315"></div>
        <div class="form-group"><label>Chủ công ty / Người nộp rút quỹ (gán 1111, cách nhau dấu phẩy):</label><input type="text" id="newCompOwners" placeholder="VD: NGUYEN VAN A, TRAN THI B"></div>
        <div class="modal-footer">
            <button class="btn btn-soft btn-sm" onclick="closeModal('modalAddCompany')">Hủy</button>
            <button class="btn btn-primary btn-sm" onclick="saveNewCompany()">💾 Lưu Công ty</button>
        </div>
    </div>
</div>

<!-- MODAL 2: THEM TK NGAN HANG -->
<div id="modalAddAccount" class="modal-backdrop">
    <div class="modal-box">
        <div class="modal-title">🏦 Thêm Tài khoản Ngân hàng</div>
        <div style="font-size: 13px; color: #5b6982; margin-bottom: 12px;">Công ty: <b id="accModalCompName"></b></div>
        <div class="form-group">
            <label>Ngân hàng:</label>
            <select id="addAccBank">
                <option value="Vietinbank">Vietinbank</option><option value="BIDV">BIDV</option><option value="LPBank">LPBank</option>
                <option value="ACB">ACB</option><option value="MB">MBBank</option><option value="Agribank">Agribank</option>
                <option value="Sacombank">Sacombank</option><option value="Vietcombank">Vietcombank</option><option value="VPBank">VPBank</option>
                <option value="MSB">MSB</option><option value="NCB">NCB</option><option value="Co-op Bank">Co-op Bank</option><option value="Khác">Khác</option>
            </select>
        </div>
        <div class="form-group"><label>Số tài khoản (*):</label><input type="text" id="addAccNumber" placeholder="Nhập số tài khoản ngân hàng..."></div>
        <div class="modal-footer">
            <button class="btn btn-soft btn-sm" onclick="closeModal('modalAddAccount')">Hủy</button>
            <button class="btn btn-primary btn-sm" onclick="saveNewAccount()">💾 Lưu Số TK</button>
        </div>
    </div>
</div>

<!-- MODAL 3: XEM / XOA DANH SACH KH / NCC -->
<div id="modalViewPartners" class="modal-backdrop">
    <div class="modal-box wide">
        <div style="display:flex; justify-content:space-between; align-items:center; gap:10px; margin-bottom:12px;">
            <div class="modal-title" id="viewPartnersTitle" style="margin:0;"></div>
            <input type="text" id="viewPartnersSearch" style="width:240px; padding:7px 10px;" placeholder="🔎 Tìm mã hoặc tên..." oninput="renderPartnerModalRows()">
        </div>
        <div class="scroll-box" style="max-height:380px;">
            <table>
                <thead>
                    <tr><th class="text-center" style="width:50px;">STT</th><th style="width:150px;">Mã đối tượng</th><th>Tên Khách hàng / NCC</th><th style="width:130px;">Mã số thuế</th><th class="text-center" style="width:65px;">Xóa</th></tr>
                </thead>
                <tbody id="viewPartnersBody"></tbody>
            </table>
        </div>
        <div class="modal-footer"><button class="btn btn-primary btn-sm" onclick="closeModal('modalViewPartners')">Đóng</button></div>
    </div>
</div>

<!-- MODAL 4: LUU QUY LUAT TU DONG (SUA 1 LAN, LAN SAU TU THEO) -->
<div id="modalSaveRule" class="modal-backdrop">
    <div class="modal-box">
        <div class="modal-title">💾 Lưu Quy Luật Tự Động Cho Công Ty [<span id="ruleCompBadge"></span>]</div>
        <div style="font-size:12.5px; color:#5b6982; background:#f5f7fd; padding:10px; border-radius:8px; margin-bottom:12px;" id="ruleRowContext"></div>
        <div class="form-group">
            <label>Từ khóa nhận diện (tên người/công ty hoặc cụm từ trong diễn giải/đối ứng):</label>
            <input type="text" id="ruleKeyword" placeholder="VD: TRAN VAN TRUONG hoặc thinh phat ceramic">
        </div>
        <div class="form-group">
            <label>Áp dụng khi:</label>
            <select id="ruleDirection">
                <option value="in">Chỉ tiền VÀO (Có > 0)</option>
                <option value="out">Chỉ tiền RA (Nợ > 0)</option>
                <option value="both">Cả tiền VÀO &amp; RA</option>
            </select>
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
            <div class="form-group"><label>Gán TK đối ứng:</label><input type="text" id="ruleTK" placeholder="VD: 1111, 131, 331, 6422, 515 hoặc trống"></div>
            <div class="form-group"><label>Gán Mã đối tượng:</label><input type="text" id="ruleMaDT" list="partnerDatalist" placeholder="VD: thinhphat ceramic (hoặc trống)"></div>
        </div>
        <div class="form-group" style="margin-top:6px;">
            <label style="display:flex; align-items:center; gap:8px; cursor:pointer; font-weight:600; color:#8c5214;">
                <input type="checkbox" id="ruleAlsoOwner" style="width:auto;" onchange="onToggleOwnerRule()">
                👑 Đây là Chủ công ty / Thủ quỹ (tự động gán TK 1111 &amp; thêm vào danh sách Chủ CT)
            </label>
        </div>
        <div class="modal-footer">
            <button class="btn btn-soft btn-sm" onclick="closeModal('modalSaveRule')">Hủy</button>
            <button class="btn btn-asp btn-sm" onclick="confirmSaveRule()">✅ Lưu &amp; Áp dụng tất cả dòng giống vậy</button>
        </div>
    </div>
</div>

<!-- MODAL 5: XEM QUY LUAT DA HOC -->
<div id="modalViewRules" class="modal-backdrop">
    <div class="modal-box wide">
        <div class="modal-title" id="viewRulesTitle">📜 Danh sách Quy luật đã lưu</div>
        <div class="scroll-box" style="max-height:380px;">
            <table>
                <thead>
                    <tr><th class="text-center" style="width:45px;">STT</th><th>Từ khóa nhận diện</th><th style="width:115px;">Chiều tiền</th><th class="text-center" style="width:95px;">TK ĐƯ</th><th style="width:160px;">Mã đối tượng</th><th class="text-center" style="width:60px;">Xóa</th></tr>
                </thead>
                <tbody id="viewRulesBody"></tbody>
            </table>
        </div>
        <div class="modal-footer"><button class="btn btn-primary btn-sm" onclick="closeModal('modalViewRules')">Đóng</button></div>
    </div>
</div>

<script>
let sharedDB = { companies: {} };
let selectedCompanyCode = '';
let allSheetsData = [];
let activeIdx = 0;
let previewMode = 'asp';
let viewingRoleKey = 'customers';
let ruleTargetCompCode = '';

async function loadSharedDB() {
    try {
        const res = await fetch('/api/partners');
        const data = await res.json();
        sharedDB = data.db || { companies: {} };
        if (!selectedCompanyCode && sharedDB.companies['lanhien']) {
            selectedCompanyCode = 'lanhien';
        }
        renderCompanyDropdown();
    } catch (e) { console.error(e); }
}

function renderCompanyDropdown() {
    const sel = document.getElementById('companySelect');
    sel.innerHTML = '<option value="">-- Tự động nhận diện theo Số tài khoản --</option>';
    Object.values(sharedDB.companies || {})
        .sort((a, b) => (a.code || '').localeCompare(b.code || ''))
        .forEach(c => {
            let opt = document.createElement('option');
            opt.value = c.code;
            opt.text = `[${c.code}] ${c.name}`;
            sel.appendChild(opt);
        });
    sel.value = (selectedCompanyCode && sharedDB.companies[selectedCompanyCode]) ? selectedCompanyCode : '';
    updateCompanyUI();
}

function updateCompanyUI() {
    const comp = sharedDB.companies[selectedCompanyCode];
    const khEl = document.getElementById('khStatus');
    const nccEl = document.getElementById('nccStatus');
    const accBox = document.getElementById('accChipsBox');
    const ownBox = document.getElementById('ownerChipsBox');
    const rulesBtn = document.getElementById('rulesBtn');

    if (!comp) {
        khEl.innerText = '0 Khách hàng'; khEl.className = 'badge-count badge-empty';
        nccEl.innerText = '0 Nhà cung cấp'; nccEl.className = 'badge-count badge-empty';
        rulesBtn.innerText = '📜 Quy luật đã học (0)';
        accBox.innerHTML = '<span style="font-size:12px;color:#7c879c;">💡 Đang ở chế độ tự nhận diện Công ty theo Số TK của từng file sao kê tải lên.</span>';
        ownBox.innerHTML = '';
        return;
    }

    const khCnt = (comp.customers || []).length;
    const nccCnt = (comp.suppliers || []).length;
    const ruleCnt = (comp.rules || []).length;
    khEl.innerText = `${khCnt} Khách hàng`; khEl.className = khCnt > 0 ? 'badge-count' : 'badge-count badge-empty';
    nccEl.innerText = `${nccCnt} Nhà cung cấp`; nccEl.className = nccCnt > 0 ? 'badge-count' : 'badge-count badge-empty';
    rulesBtn.innerText = `📜 Quy luật đã học (${ruleCnt})`;

    const accs = comp.accounts || [];
    accBox.innerHTML = '<span style="font-size:12px;font-weight:600;color:#5b6982;">🏦 Số TK:</span> ' +
        (accs.length ? accs.map((a, idx) =>
            `<span class="acc-chip">${a.bank}: <b>${a.account}</b> <button class="chip-del" onclick="removeAccount(${idx})">×</button></span>`
        ).join('') : '<span style="font-size:12px;color:#8892a6;">Chưa gắn số TK.</span>');

    const owners = comp.owners || [];
    ownBox.innerHTML = '<span style="font-size:12px;font-weight:600;color:#8c5214;">👑 Chủ CT (tự gán 1111):</span> ' +
        owners.map((o, idx) =>
            `<span class="owner-chip">${o} <button class="chip-del" onclick="removeOwner(${idx})">×</button></span>`
        ).join('') +
        ' <button class="btn btn-amber btn-xs" onclick="addOwnerPrompt()">+ Thêm Chủ CT</button>';
}

function onSelectCompany() {
    selectedCompanyCode = document.getElementById('companySelect').value;
    updateCompanyUI();
}

async function persistDBToServer() {
    const res = await fetch('/api/partners/save', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(sharedDB)
    });
    const data = await res.json();
    document.getElementById('syncBadge').innerText = data.message || '✅ Đã lưu';
}

function closeModal(id) { document.getElementById(id).style.display = 'none'; }

function openAddCompanyModal() {
    document.getElementById('newCompName').value = '';
    document.getElementById('newCompCode').value = '';
    document.getElementById('newCompAcc').value = '';
    document.getElementById('newCompOwners').value = '';
    document.getElementById('modalAddCompany').style.display = 'flex';
}

async function saveNewCompany() {
    const name = document.getElementById('newCompName').value.trim();
    const code = document.getElementById('newCompCode').value.trim().toLowerCase().replace(/\s+/g, '');
    const bank = document.getElementById('newCompBank').value;
    const acc = document.getElementById('newCompAcc').value.trim();
    const ownersRaw = document.getElementById('newCompOwners').value;
    const owners = ownersRaw.split(',').map(x => x.trim()).filter(Boolean);

    if (!name || !code) { alert('Vui lòng nhập đầy đủ Tên công ty và Mã công ty!'); return; }
    if (!sharedDB.companies[code]) {
        sharedDB.companies[code] = { code, name, accounts: [], owners: [], customers: [], suppliers: [], rules: [] };
    } else {
        sharedDB.companies[code].name = name;
    }
    if (acc && !sharedDB.companies[code].accounts.some(a => a.account === acc)) {
        sharedDB.companies[code].accounts.push({ bank, account: acc });
    }
    owners.forEach(o => {
        if (!sharedDB.companies[code].owners.includes(o)) sharedDB.companies[code].owners.push(o);
    });
    selectedCompanyCode = code;
    closeModal('modalAddCompany');
    renderCompanyDropdown();
    await persistDBToServer();
}

function openAddAccountModal() {
    if (!selectedCompanyCode || !sharedDB.companies[selectedCompanyCode]) {
        alert('Hãy chọn một công ty ở ô phía trên trước!'); return;
    }
    document.getElementById('accModalCompName').innerText = `[${selectedCompanyCode}] ${sharedDB.companies[selectedCompanyCode].name}`;
    document.getElementById('addAccNumber').value = '';
    document.getElementById('modalAddAccount').style.display = 'flex';
}

async function saveNewAccount() {
    const bank = document.getElementById('addAccBank').value;
    const acc = document.getElementById('addAccNumber').value.trim();
    if (!acc) { alert('Vui lòng nhập Số tài khoản!'); return; }
    const comp = sharedDB.companies[selectedCompanyCode];
    if (!comp.accounts) comp.accounts = [];
    if (!comp.accounts.some(a => a.account === acc)) comp.accounts.push({ bank, account: acc });
    closeModal('modalAddAccount');
    updateCompanyUI();
    await persistDBToServer();
}

async function removeAccount(idx) {
    const comp = sharedDB.companies[selectedCompanyCode];
    if (!comp || !comp.accounts) return;
    if (!confirm(`Gỡ số TK ${comp.accounts[idx].account}?`)) return;
    comp.accounts.splice(idx, 1);
    updateCompanyUI();
    await persistDBToServer();
}

async function addOwnerPrompt() {
    if (!selectedCompanyCode) return;
    const name = prompt('Nhập tên Chủ công ty / Thủ quỹ (VD: TRAN VAN TRUONG):');
    if (!name || !name.trim()) return;
    const comp = sharedDB.companies[selectedCompanyCode];
    if (!comp.owners) comp.owners = [];
    if (!comp.owners.includes(name.trim())) comp.owners.push(name.trim());
    updateCompanyUI();
    await persistDBToServer();
}

async function removeOwner(idx) {
    const comp = sharedDB.companies[selectedCompanyCode];
    if (!comp || !comp.owners) return;
    comp.owners.splice(idx, 1);
    updateCompanyUI();
    await persistDBToServer();
}

async function deleteCompany() {
    if (!selectedCompanyCode) { alert('Hãy chọn một công ty để xóa!'); return; }
    if (!confirm(`Xóa toàn bộ danh bạ của [${selectedCompanyCode}]?`)) return;
    delete sharedDB.companies[selectedCompanyCode];
    selectedCompanyCode = '';
    renderCompanyDropdown();
    await persistDBToServer();
}

function triggerPartnerUpload(inputId) {
    if (!selectedCompanyCode) { alert('Vui lòng chọn Công ty ở ô phía trên trước!'); return; }
    document.getElementById(inputId).click();
}

async function uploadPartnerList(roleKey, actionMode, inputEl) {
    if (!inputEl.files.length || !selectedCompanyCode) return;
    const fd = new FormData();
    fd.append('file', inputEl.files[0]);
    try {
        const res = await fetch('/parse_partners', { method: 'POST', body: fd });
        const data = await res.json();
        if (data.error) { alert(data.error); return; }

        const incoming = data.items || [];
        const label = roleKey === 'customers' ? 'Khách hàng' : 'Nhà cung cấp';

        if (actionMode === 'replace') {
            sharedDB.companies[selectedCompanyCode][roleKey] = incoming;
            updateCompanyUI();
            await persistDBToServer();
            alert(`✅ Đã nạp mới ${incoming.length} ${label} cho [${selectedCompanyCode}]!`);
        } else {
            const current = sharedDB.companies[selectedCompanyCode][roleKey] || [];
            const mapByCode = new Map();
            current.forEach((item, idx) => mapByCode.set(item.code.toLowerCase(), idx));
            let added = 0, updated = 0;
            incoming.forEach(item => {
                const k = item.code.toLowerCase();
                if (mapByCode.has(k)) { current[mapByCode.get(k)] = item; updated++; }
                else { mapByCode.set(k, current.length); current.push(item); added++; }
            });
            sharedDB.companies[selectedCompanyCode][roleKey] = current;
            updateCompanyUI();
            await persistDBToServer();
            alert(`🔄 Đã Update DS ${label} cho [${selectedCompanyCode}]: Thêm ${added}, cập nhật ${updated} (Tổng: ${current.length})!`);
        }
    } catch (e) {
        alert('Lỗi đọc file danh sách: ' + e);
    } finally { inputEl.value = ''; }
}

function openViewPartnersModal(roleKey) {
    if (!selectedCompanyCode || !sharedDB.companies[selectedCompanyCode]) { alert('Chọn một công ty để xem!'); return; }
    viewingRoleKey = roleKey;
    document.getElementById('viewPartnersTitle').innerText = `📋 Danh sách ${roleKey === 'customers' ? 'Khách hàng' : 'Nhà cung cấp'} - [${selectedCompanyCode}]`;
    document.getElementById('viewPartnersSearch').value = '';
    renderPartnerModalRows();
    document.getElementById('modalViewPartners').style.display = 'flex';
}

function renderPartnerModalRows() {
    const comp = sharedDB.companies[selectedCompanyCode];
    const list = (comp && comp[viewingRoleKey]) ? comp[viewingRoleKey] : [];
    const kw = document.getElementById('viewPartnersSearch').value.trim().toLowerCase();
    const tbody = document.getElementById('viewPartnersBody');
    tbody.innerHTML = '';
    const filtered = list.map((item, idx) => ({ ...item, _idx: idx }))
        .filter(it => !kw || it.code.toLowerCase().includes(kw) || it.name.toLowerCase().includes(kw) || (it.mst || '').includes(kw));
    if (!filtered.length) {
        tbody.innerHTML = '<tr><td colspan="5" class="text-center" style="color:#888;padding:20px;">Chưa có dữ liệu</td></tr>';
        return;
    }
    filtered.forEach((it, i) => {
        let tr = document.createElement('tr');
        tr.innerHTML = `<td class="text-center">${i + 1}</td><td><b>${it.code}</b></td><td>${it.name}</td><td>${it.mst || ''}</td>
            <td class="text-center"><button class="btn btn-peach btn-xs" onclick="deletePartnerItem(${it._idx})">×</button></td>`;
        tbody.appendChild(tr);
    });
}

async function deletePartnerItem(origIdx) {
    sharedDB.companies[selectedCompanyCode][viewingRoleKey].splice(origIdx, 1);
    renderPartnerModalRows(); updateCompanyUI(); await persistDBToServer();
}

function openViewRulesModal() {
    if (!selectedCompanyCode || !sharedDB.companies[selectedCompanyCode]) { alert('Chọn một công ty để xem!'); return; }
    document.getElementById('viewRulesTitle').innerText = `📜 Quy luật đã học - [${selectedCompanyCode}]`;
    renderRulesModalRows();
    document.getElementById('modalViewRules').style.display = 'flex';
}

function renderRulesModalRows() {
    const comp = sharedDB.companies[selectedCompanyCode];
    const rules = (comp && comp.rules) ? comp.rules : [];
    const tbody = document.getElementById('viewRulesBody');
    tbody.innerHTML = '';
    if (!rules.length) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center" style="color:#888;padding:20px;">Chưa có quy luật nào được lưu</td></tr>';
        return;
    }
    const dirLabel = { in: 'Tiền VÀO (Có)', out: 'Tiền RA (Nợ)', both: 'Cả hai chiều' };
    rules.forEach((r, idx) => {
        let tr = document.createElement('tr');
        tr.innerHTML = `<td class="text-center">${idx + 1}</td><td><b>${r.keyword}</b></td><td>${dirLabel[r.direction] || r.direction}</td>
            <td class="text-center tk-badge">${r.tk || '(trống)'}</td><td>${r.madt || ''}</td>
            <td class="text-center"><button class="btn btn-peach btn-xs" onclick="deleteRuleItem(${idx})">×</button></td>`;
        tbody.appendChild(tr);
    });
}

async function deleteRuleItem(idx) {
    sharedDB.companies[selectedCompanyCode].rules.splice(idx, 1);
    renderRulesModalRows(); updateCompanyUI(); await persistDBToServer();
}

function showFiles() {
    const inp = document.getElementById('files');
    const names = Array.from(inp.files).map(f => f.name);
    document.getElementById('fileList').innerText = names.length ? `Đã chọn ${names.length} file: ${names.join(', ')}` : '';
}

async function uploadAndProcess() {
    const inp = document.getElementById('files');
    if (!inp.files.length) { alert('Vui lòng chọn ít nhất 1 file sao kê!'); return; }

    const formData = new FormData();
    for (let f of inp.files) formData.append('files', f);
    formData.append('password', document.getElementById('pwd').value);
    formData.append('fallback_company_code', selectedCompanyCode);

    document.getElementById('loader').style.display = 'block';
    document.getElementById('resultBox').style.display = 'none';

    try {
        const res = await fetch('/process', { method: 'POST', body: formData });
        const data = await res.json();
        if (data.error) { alert(data.error); return; }

        const tbody = document.getElementById('summaryBody');
        tbody.innerHTML = '';
        data.summary.forEach(r => {
            let tr = document.createElement('tr');
            let cls = r.status === 'Khớp chuẩn' ? 'ok' : 'warn';
            tr.innerHTML = `<td>${r.file}</td><td><b>${r.bank}</b></td><td>${r.account}</td>
                <td class="text-right">${r.count}</td><td class="text-right debit">${r.debit}</td>
                <td class="text-right credit">${r.credit}</td><td><span class="${cls}">${r.status}</span></td>`;
            tbody.appendChild(tr);
        });

        allSheetsData = data.sheets || [];
        renderTabs();
        document.getElementById('resultBox').style.display = 'block';
    } catch (e) {
        alert('Có lỗi xảy ra: ' + e);
    } finally {
        document.getElementById('loader').style.display = 'none';
    }
}

function switchPreviewMode(mode) {
    previewMode = mode;
    document.getElementById('btnMode4').classList.toggle('active', mode === '4col');
    document.getElementById('btnModeAsp').classList.toggle('active', mode === 'asp');
    filterPreview();
}

function renderTabs() {
    const tabsEl = document.getElementById('tabs');
    tabsEl.innerHTML = '';
    allSheetsData.forEach((s, idx) => {
        let btn = document.createElement('button');
        btn.className = 'tab-btn' + (idx === 0 ? ' active' : '');
        btn.innerText = `${s.file} (${s.rows.length})`;
        btn.onclick = () => selectTab(idx);
        tabsEl.appendChild(btn);
    });
    if (allSheetsData.length > 0) selectTab(0);
}

function populatePartnerDatalist(compCode) {
    const dl = document.getElementById('partnerDatalist');
    dl.innerHTML = '';
    const comp = sharedDB.companies[compCode];
    if (!comp) return;
    const seen = new Set();
    [...(comp.customers || []), ...(comp.suppliers || [])].forEach(it => {
        if (!seen.has(it.code)) {
            seen.add(it.code);
            let opt = document.createElement('option');
            opt.value = it.code;
            opt.label = it.name;
            dl.appendChild(opt);
        }
    });
}

function selectTab(idx) {
    activeIdx = idx;
    document.querySelectorAll('.tab-btn').forEach((b, i) => b.classList.toggle('active', i === idx));
    document.getElementById('searchBox').value = '';
    const s = allSheetsData[idx];
    populatePartnerDatalist(s.company_code || selectedCompanyCode);
    document.getElementById('sheetMeta').innerHTML =
        `Ngân hàng: <b>${s.bank}</b> • Số TK: <b>${s.account || '—'}</b> • Công ty: <b>[${s.company_code || 'chưa gắn'}] ${s.customer || ''}</b> • Giao dịch: <b>${s.rows.length}</b>`;
    renderPreviewRows(s.rows);
}

function filterPreview() {
    const kw = document.getElementById('searchBox').value.trim().toLowerCase();
    const rows = allSheetsData[activeIdx]?.rows || [];
    if (!kw) { renderPreviewRows(rows); return; }
    const filtered = rows.filter(r =>
        (r.Ngay || '').toLowerCase().includes(kw) ||
        (r.MoTa || '').toLowerCase().includes(kw) ||
        (r.DoiTac || '').toLowerCase().includes(kw) ||
        (r.No || '').toLowerCase().includes(kw) ||
        (r.Co || '').toLowerCase().includes(kw) ||
        (r.TK || '').toLowerCase().includes(kw) ||
        (r.MaDT || '').toLowerCase().includes(kw)
    );
    renderPreviewRows(filtered);
}

function updateRowField(rowObjIndex, field, value) {
    if (allSheetsData[activeIdx] && allSheetsData[activeIdx].rows[rowObjIndex]) {
        allSheetsData[activeIdx].rows[rowObjIndex][field] = value.trim();
    }
}

function renderPreviewRows(rows) {
    const thead = document.getElementById('previewHead');
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';

    if (previewMode === 'asp') {
        thead.innerHTML = `<tr>
            <th class="text-center" style="width: 100px;">Ngày</th>
            <th class="text-right" style="width: 120px;">Nợ</th>
            <th class="text-right" style="width: 120px;">Có</th>
            <th>Diễn giải</th>
            <th class="text-center" style="width: 95px;">TK ĐƯ</th>
            <th style="width: 150px;">Mã đối tượng</th>
            <th class="text-center" style="width: 90px;">Quy luật</th>
        </tr>`;
    } else {
        thead.innerHTML = `<tr>
            <th class="text-center" style="width: 115px;">Ngày</th>
            <th class="text-right" style="width: 145px;">Nợ</th>
            <th class="text-right" style="width: 145px;">Có</th>
            <th>Diễn giải</th>
        </tr>`;
    }

    const colCount = previewMode === 'asp' ? 7 : 4;
    if (!rows.length) {
        tbody.innerHTML = `<tr><td colspan="${colCount}" class="text-center" style="color:#888;padding:24px;">Không có dữ liệu giao dịch phù hợp</td></tr>`;
        return;
    }

    const sourceRows = allSheetsData[activeIdx].rows;
    rows.forEach(r => {
        const origIdx = sourceRows.indexOf(r);
        let tr = document.createElement('tr');
        let noDisp = (r.No && r.No !== '0') ? r.No : '';
        let coDisp = (r.Co && r.Co !== '0') ? r.Co : '';

        if (previewMode === 'asp') {
            tr.innerHTML = `
                <td class="text-center"><b>${r.Ngay}</b></td>
                <td class="text-right debit">${noDisp}</td>
                <td class="text-right credit">${coDisp}</td>
                <td>${r.MoTa}${r.DoiTac ? `<div style="font-size:11.5px;color:#7a8599;margin-top:2px;">Đối ứng: ${r.DoiTac}</div>` : ''}</td>
                <td><input class="cell-input tk-badge" value="${r.TK || ''}" onchange="updateRowField(${origIdx}, 'TK', this.value)" placeholder="Trống"></td>
                <td><input class="cell-input" list="partnerDatalist" value="${r.MaDT || ''}" onchange="updateRowField(${origIdx}, 'MaDT', this.value)" placeholder="Trống"></td>
                <td class="text-center"><button class="btn btn-soft btn-xs" onclick="openSaveRuleModal(${origIdx})" title="Lưu quy luật cho các dòng giống vậy và các lần sau">💾 Lưu</button></td>
            `;
        } else {
            tr.innerHTML = `
                <td class="text-center"><b>${r.Ngay}</b></td>
                <td class="text-right debit">${noDisp}</td>
                <td class="text-right credit">${coDisp}</td>
                <td>${r.MoTa}</td>
            `;
        }
        tbody.appendChild(tr);
    });
}

function suggestRuleKeyword(row) {
    // Uu tien goi y tu cot DoiTac (tru tien mat)
    let dt = (row.DoiTac || '').split('–')[0].split('- A/C')[0].split('A/C:')[0].trim();
    if (dt && !dt.toLowerCase().includes('tien mat') && dt.length >= 3) return dt;
    // Neu khong co DoiTac, lam sach chuoi MoTa de goi y ten nguoi / cong ty
    let m = (row.MoTa || '').replace(/CT nhanh 247 (den|di):/i, '').replace(/chuyen tien|nop tien|rut tien/gi, '').trim();
    return m.slice(0, 45).trim() || row.MoTa;
}

function openSaveRuleModal(origIdx) {
    const s = allSheetsData[activeIdx];
    const row = s.rows[origIdx];
    ruleTargetCompCode = s.company_code || selectedCompanyCode;
    if (!ruleTargetCompCode || !sharedDB.companies[ruleTargetCompCode]) {
        alert('File sao kê này chưa gắn với Công ty nào. Vui lòng chọn Công ty ở ô phía trên trước!');
        return;
    }
    const isCo = row.Co && row.Co !== '0';
    document.getElementById('ruleCompBadge').innerText = ruleTargetCompCode;
    document.getElementById('ruleRowContext').innerHTML = `<b>Diễn giải:</b> ${row.MoTa}` + (row.DoiTac ? `<br><b>Đối ứng:</b> ${row.DoiTac}` : '');
    document.getElementById('ruleKeyword').value = suggestRuleKeyword(row);
    document.getElementById('ruleDirection').value = isCo ? 'in' : 'out';
    document.getElementById('ruleTK').value = row.TK || '';
    document.getElementById('ruleMaDT').value = row.MaDT || '';
    document.getElementById('ruleAlsoOwner').checked = false;
    document.getElementById('modalSaveRule').style.display = 'flex';
}

function onToggleOwnerRule() {
    if (document.getElementById('ruleAlsoOwner').checked) {
        document.getElementById('ruleTK').value = '1111';
        document.getElementById('ruleMaDT').value = '';
    }
}

function normalizeForMatch(s) {
    return (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
}

async function confirmSaveRule() {
    const kw = document.getElementById('ruleKeyword').value.trim();
    const dir = document.getElementById('ruleDirection').value;
    const tk = document.getElementById('ruleTK').value.trim();
    const madt = document.getElementById('ruleMaDT').value.trim();
    const isOwner = document.getElementById('ruleAlsoOwner').checked;

    if (!kw) { alert('Vui lòng nhập Từ khóa nhận diện!'); return; }
    const comp = sharedDB.companies[ruleTargetCompCode];

    if (isOwner) {
        if (!comp.owners) comp.owners = [];
        if (!comp.owners.includes(kw)) comp.owners.push(kw);
    } else {
        if (!comp.rules) comp.rules = [];
        // Cap nhat neu da co rule cung keyword + direction
        const existingIdx = comp.rules.findIndex(r => r.keyword.toLowerCase() === kw.toLowerCase() && r.direction === dir);
        if (existingIdx !== -1) comp.rules[existingIdx] = { keyword: kw, direction: dir, tk, madt };
        else comp.rules.push({ keyword: kw, direction: dir, tk, madt });
    }

    // Ap dung ngay cho TAT CA cac dong khop trong sheet hien tai
    const kwNorm = normalizeForMatch(kw);
    let updatedCount = 0;
    allSheetsData.forEach(sheet => {
        if ((sheet.company_code || selectedCompanyCode) !== ruleTargetCompCode) return;
        sheet.rows.forEach(r => {
            const rowText = normalizeForMatch((r.MoTa || '') + ' ' + (r.DoiTac || ''));
            const rowDir = (r.Co && r.Co !== '0') ? 'in' : 'out';
            if (rowText.includes(kwNorm) && (isOwner || dir === 'both' || dir === rowDir)) {
                r.TK = isOwner ? '1111' : tk;
                r.MaDT = isOwner ? '' : madt;
                updatedCount++;
            }
        });
    });

    closeModal('modalSaveRule');
    updateCompanyUI();
    filterPreview();
    await persistDBToServer();
    alert(`✅ Đã lưu quy luật "${kw}" và tự động cập nhật ${updatedCount} dòng trên bảng! Các lần sau sẽ tự động áp dụng.`);
}

async function downloadExcel(mode) {
    if (!allSheetsData.length) return;
    try {
        const res = await fetch('/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ mode: mode, sheets: allSheetsData })
        });
        if (!res.ok) { alert('Lỗi khi tạo file Excel'); return; }
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = mode === 'asp' ? 'KetXuat_Cho_ASP.xlsx' : 'SaoKe_4Cot_Chuan.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    } catch (e) {
        alert('Lỗi tải file: ' + e);
    }
}

loadSharedDB();
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML_UI)


@app.route("/api/partners", methods=["GET"])
def api_get_partners():
    return jsonify({"db": load_shared_db()})


@app.route("/api/partners/save", methods=["POST"])
def api_save_partners():
    payload = request.get_json(silent=True) or {}
    ok, msg = save_shared_db(payload)
    return jsonify({"ok": ok, "message": msg})


@app.route("/parse_partners", methods=["POST"])
def parse_partners_route():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "Không có file danh sách được tải lên"}), 400
    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            p = os.path.join(tmp_dir, os.path.basename(f.filename))
            f.save(p)
            items = parse_partner_excel(p)
        return jsonify({"items": items})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/process", methods=["POST"])
def process_files():
    uploaded = request.files.getlist("files")
    pwd = request.form.get("password", "").strip() or None
    fallback_comp_code = request.form.get("fallback_company_code", "").strip()
    shared_db = load_shared_db()
    companies_map = shared_db.get("companies", {})

    acc_to_comp = {}
    for c_code, c_obj in companies_map.items():
        for acc_item in c_obj.get("accounts", []):
            acc_num = str(acc_item.get("account", "")).strip()
            if acc_num:
                acc_to_comp[acc_num] = c_code

    if not uploaded:
        return jsonify({"error": "Không có file nào được tải lên"}), 400

    with tempfile.TemporaryDirectory() as tmp_dir:
        paths = []
        for f in uploaded:
            p = os.path.join(tmp_dir, os.path.basename(f.filename))
            f.save(p)
            paths.append(p)

        summary, sheets = [], []
        for p in paths:
            res = extract_any(p, password=pwd)
            warns = [] if res.get('loi') else validate(res)
            hi = res.get('header_info', {})
            acc_no = str(hi.get('account_no', '') or '').strip()

            comp_code = acc_to_comp.get(acc_no, "") or fallback_comp_code
            comp_obj = companies_map.get(comp_code, {})
            owner_name = hi.get('customer') or hi.get('account_name') or comp_obj.get('name', '')
            customers = comp_obj.get("customers", [])
            suppliers = comp_obj.get("suppliers", [])
            owners = comp_obj.get("owners", [])
            rules = comp_obj.get("rules", [])
            bank_code = _resolve_bank_partner_code(res.get('bank'), res.get('source_file', ''))

            t_deb = to_num(res['total_debit']) if res.get('total_debit') else sum(to_num(x['No']) for x in res['rows'])
            t_cred = to_num(res['total_credit']) if res.get('total_credit') else sum(to_num(x['Co']) for x in res['rows'])

            if res.get('loi'): status = "❌ LỖI: " + res['loi']
            elif warns: status = "⚠️ CẢNH BÁO: " + " | ".join(warns[:2])
            else: status = "Khớp chuẩn"

            summary.append({
                "file": res['source_file'],
                "bank": str(res['bank']).upper(),
                "account": acc_no,
                "count": len(res['rows']),
                "debit": f"{t_deb:,.0f}",
                "credit": f"{t_cred:,.0f}",
                "status": status
            })

            preview_rows = []
            for r in res['rows']:
                no_v = to_num(r['No'])
                co_v = to_num(r['Co'])
                tk_doi_ung, ma_dt = classify_for_asp(
                    r, owner_name=owner_name, owner_code=comp_code, owner_acc=acc_no,
                    customers=customers, suppliers=suppliers, owners=owners, rules=rules,
                    bank_code=bank_code, acc_to_comp=acc_to_comp
                )
                preview_rows.append({
                    "Ngay": r['Ngay'],
                    "No": f"{no_v:,.0f}" if no_v != 0 else "0",
                    "Co": f"{co_v:,.0f}" if co_v != 0 else "0",
                    "MoTa": r['MoTa'],
                    "DoiTac": r.get('DoiTac', ''),
                    "TK": tk_doi_ung,
                    "MaDT": ma_dt
                })

            sheets.append({
                "file": res['source_file'],
                "bank": str(res['bank']).upper(),
                "account": acc_no,
                "customer": owner_name,
                "company_code": comp_code,
                "error": res.get('loi'),
                "rows": preview_rows
            })

    return jsonify({"summary": summary, "sheets": sheets})


@app.route("/export", methods=["POST"])
def export_excel_route():
    payload = request.get_json(silent=True) or {}
    mode = payload.get("mode", "4col")
    sheets = payload.get("sheets", [])
    xlsx_bytes = create_workbook_from_sheets(sheets, mode=mode)
    fname = "KetXuat_Cho_ASP.xlsx" if mode == "asp" else "SaoKe_4Cot_Chuan.xlsx"
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=fname
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    if not os.environ.get("RENDER") and not os.environ.get("PYTHONANYWHERE_DOMAIN"):
        Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host="0.0.0.0", port=port)
