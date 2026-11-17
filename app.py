# -*- coding: utf-8 -*-
"""
CONG CU TRICH XUAT SAO KE NGAN HANG & KET XUAT MISA ASP / MISA SME
- PDF: Vietcombank (Sao ke co ke o, Sao ke khong ke o, va So phu Giay bao No/Co),
       VietinBank, BIDV, LPBank, ACB, MBBank
- HTML: VPBank (T24 Core)
- Excel/CSV: Sacombank, Agribank IPCAS, Excel khac
- Ho tro tai file .zip So phu (ghep vao Sao ke, tu dong TACH PHI CHUYEN TIEN VCB -> TK 6422 & Doi tuong Ngan hang,
  hoac tu tong hop thanh bang giao dich neu chi tai So phu)
- Xuat Excel: Mau 4 cot, MISA ASP (6 cot), va MISA SME (2 Sheet Thu_Tien_Gui & Chi_Tien_Gui trong cung 1 file)
- Dinh dang ngay chuan: dd/mm/yyyy
"""

import base64
import csv
import gc
from html.parser import HTMLParser
import io
import json
import os
import re
import shutil
import tempfile
import threading
import time
import unicodedata
import urllib.request
import uuid
import webbrowser
import zipfile
from threading import Timer

from flask import Flask, jsonify, render_template_string, request, send_file
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pdfplumber
import xlrd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 80 * 1024 * 1024

_JOB_ROOT = os.path.join(tempfile.gettempdir(), "saoke_render_jobs")
os.makedirs(_JOB_ROOT, exist_ok=True)


# ============================================================================
# 0. DU LIEU MAC DINH: 16 CONG TY & TAI KHOAN NGAN HANG + KHACH HANG MAU
# ============================================================================

_DEFAULT_ACCOUNTS = [
    ("CÔNG TY TNHH THƯƠNG MẠI KHÔI HƯƠNG", "khoihuong", "LPBank", "8181222666"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ XÂY DỰNG THỊNH HIỀN", "thinhhien", "LPBank", "5565568868"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ HÙNG LONG", "hunglong", "LPBank", "072710420001"),
    ("CÔNG TY TNHH THƯƠNG MẠI VÀ SẢN XUẤT MINH HUY HMA", "minhhuy", "LPBank", "8386556789"),
    ("CÔNG TY TNHH DV THƯƠNG MẠI PHÚC THÀNH", "phucthanh", "Vietcombank", "34643548S01"),
    ("CÔNG TY TNHH DV THƯƠNG MẠI PHÚC THÀNH", "phucthanh", "MSB", "83988888"),
    ("CÔNG TY TNHH NANO HOÀNG ANH", "hoanganh", "LPBank", "629899999"),
    ("CÔNG TY TNHH DỊCH VỤ THƯƠNG MẠI HÀ KỀN", "haken", "Vietcombank", "1066158986"),
    ("CÔNG TY TNHH XĂNG DẦU THÀNH ANH", "thanhanh", "Vietinbank", "116000149017"),
    ("CÔNG TY TNHH TMDV LỘC HOÀNG", "lochoang", "BIDV", "8600065578"),
    ("CÔNG TY TRÁCH NHIỆM HỮU HẠN LÂN HIỀN", "lanhien", "BIDV", "4820180367"),
    ("CÔNG TY TRÁCH NHIỆM HỮU HẠN LÂN HIỀN", "lanhien", "Vietinbank", "111000073812"),
    ("CÔNG TY TNHH THƯƠNG MẠI THÚY KHÁNH", "thuykhanh", "Vietinbank", "118003052315"),
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
    {"code": "0700769778", "name": "CÔNG TY TNHH TƯ VẤN ĐẦU TƯ XÂY DỰNG", "mst": "0700769778"},
    {"code": "88 HÀ NAM", "name": "CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG VÀ TM AN PHÚ", "mst": "0700918003"},
    {"code": "an phú", "name": "CÔNG TY TNHH ĐẦU TƯ VÀ TM AN PHÚ", "mst": ""},
    {"code": "CNS", "name": "CÔNG TY TNHH GIẢI PHÁP CÔNG NGHỆ VÀ ĐẦU TƯ", "mst": ""},
    {"code": "cuongkhoi", "name": "CÔNG TY TNHH THƯƠNG MẠI CƯỜNG KHÔI", "mst": "0700318528"},
    {"code": "DABACO", "name": "CÔNG TY TNHH LỢN GIỐNG DABACO HÀ NAM", "mst": "0700757934"},
    {"code": "DG", "name": "Công Ty TNHH ĐG Trung Kiên", "mst": "0700617302"},
    {"code": "ducthien", "name": "CÔNG TY TNHH TƯ VẤN VÀ ĐẦU TƯ XÂY DỰNG ĐỨC THIỆN", "mst": "0700614397"},
    {"code": "gialinh", "name": "CÔNG TY TNHH ĐẦU TƯ GIA LINH", "mst": "0700880494"},
    {"code": "hainam", "name": "CÔNG TY CỔ PHẦN TƯ VẤN ĐTXD HẢI NAM", "mst": "0600447544"},
    {"code": "HDT", "name": "Công ty Cổ phần tập đoàn đầu tư xây dựng và phát triển nhà HDT", "mst": "0101585316"},
    {"code": "hkdnguyenvananh", "name": "HỘ KINH DOANH NGUYỄN VĂN ANH - 1999", "mst": "036099011164"},
    {"code": "hoangha", "name": "CÔNG TY TNHH ĐẦU TƯ HỒNG PHÚ", "mst": "0700776292"},
    {"code": "hungbao", "name": "CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG VÀ DV TM KHÁNH HƯNG", "mst": "0700690334"},
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
    {"code": "tmhanam", "name": "CÔNG TY CP DỊCH VỤ THƯƠNG MẠI HÀ NAM", "mst": "0700217336"},
    {"code": "truonghunghn", "name": "CÔNG TY TNHH TRƯỜNG HƯNG HN", "mst": "0700898124"},
    {"code": "tuanthuHN", "name": "CÔNG TY TNHH TUẤN THU HN", "mst": "0700847377"},
    {"code": "TÙNG ANH", "name": "CÔNG TY TNHH CẤP THOÁT NƯỚC TÙNG ANH", "mst": "0700256021"},
    {"code": "vuongthin", "name": "CÔNG TY TNHH VƯƠNG THÌN", "mst": "0700921077"},
    {"code": "XÂY DỰNG 3A", "name": "CÔNG TY TNHH XÂY DỰNG 3A", "mst": "0700805458"},
    {"code": "xuankhanhHN", "name": "CÔNG TY TNHH XUÂN KHÁNH HN", "mst": ""},
]


def _build_initial_db():
    companies = {}
    for name, code, bank, acc in _DEFAULT_ACCOUNTS:
        if code not in companies:
            companies[code] = {
                "code": code, "name": name,
                "accounts": [], "owners": [], "customers": [], "suppliers": [], "rules": []
            }
        companies[code]["accounts"].append({"bank": bank, "account": acc})
    if "lanhien" in companies:
        companies["lanhien"]["customers"] = list(_DEFAULT_LANHIEN_CUSTOMERS)
        companies["lanhien"]["owners"] = ["NGUYEN VAN LAN"]
    if "thuykhanh" in companies:
        companies["thuykhanh"]["owners"] = ["TRAN VAN TRUONG", "TRAN THI THUY"]
    if "thanhanh" in companies:
        companies["thanhanh"]["owners"] = ["NGUYEN TIEN THANH"]
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


# ============================================================================
# 1. TIEN ICH CHUAN HOA NGAY THANG & TIEN TE
# ============================================================================

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
        if 1 <= d <= 12 and 1 <= m <= 31:
            return f"{m:02d}/{d:02d}/{y:04d}"
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


def _vcb_comma_dec_to_std_str(val):
    """57,750,000,00 -> 57,750,000 ; 5,545,00 -> 5,545 ; 5.500 -> 5,500"""
    s = str(val or "").strip()
    s = re.sub(r'(?i)\s*(vnd|vnđ|đ)\b', '', s).strip()
    if not s:
        return "0"
    # Dang VCB So phu: 57,750,000,00 hoac 5,500,00 (dau phay cuoi la 2 so thap phan)
    if re.match(r'^-?\d{1,3}(,\d{3})+,\d{2}$', s) or re.match(r'^-?\d+,\d{2}$', s):
        s = s[:-3].replace(",", "")
        try:
            num = abs(float(s))
            return f"{num:,.0f}" if num != 0 else "0"
        except ValueError:
            return "0"
    if re.match(r'^-?\d{1,3}(\.\d{3})+(,\d+)?$', s):
        s = s.replace(".", "").replace(",", ".")
        try:
            num = abs(float(s))
            return f"{num:,.0f}" if num != 0 else "0"
        except ValueError:
            return "0"
    return _vn_money_to_std_str(s)


# ============================================================================
# 2. NHAN DIEN NGAN HANG TU NOI DUNG FILE PDF
# ============================================================================

def detect_bank(text):
    t = unicodedata.normalize("NFC", text or "")
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
    if (
        "GIẤY BÁO NỢ - DEBIT ADVICE" in t
        or "GIẤY BÁO CÓ - CREDIT ADVICE" in t
        or "GIẤY BÁO NỢ / DEBIT ADVICE" in t
        or "GIẤY BÁO CÓ / CREDIT ADVICE" in t
        or "SỔ HẠCH TOÁN CHI TIẾT" in t
        or (("GIẤY BÁO NỢ" in t or "DEBIT ADVICE" in t) and "SAO KÊ TÀI KHOẢN" not in t.upper())
    ):
        return "vcb_sophu"
    if "www.vietcombank.com.vn" in t or "VIETCOMBANK" in t.upper():
        if "VCB - iB@nking" in t or "Nội dung chi tiết" in t or "Transaction In Detail" in t:
            return "vietcombank_vcb2"
        if "Số tham chiếu" in t or "Số tiền ghi nợ" in t:
            return "vietcombank_vcb1"
    return "unknown"


# ============================================================================
# 3A. VIETINBANK PDF
# ============================================================================

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
            page.flush_cache()
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
                'STT': anchor['text'],
                'Ngay': ngay_chuan,
                'MoTa': txt("MoTa"),
                'No': txt("No"),
                'Co': txt("Co"),
                'SoDu': txt("SoDu"),
                'DoiTac': txt("DoiTac"),
                'MaGD': txt("SoGD"),
            })
        page.flush_cache()

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# ============================================================================
# 3B. BIDV PDF
# ============================================================================

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
                if "STT" in label and "No" in label:
                    continue
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
                if len(row) < 11:
                    continue
                stt = row[0]
                if not stt or not stt.isdigit():
                    continue
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
        page.flush_cache()

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# ============================================================================
# 3C. LPBANK PDF (7 cot moi & 10 cot cu)
# ============================================================================

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
                if "Ngày giao dịch" in col0 or "Txn. date" in col0:
                    continue
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
                    'STT': str(stt_counter),
                    'Ngay': normalize_date_ddmmyyyy(row[0]),
                    'MoTa': row[2],
                    'No': no_val,
                    'Co': co_val,
                    'SoDu': sodu_val,
                    'DoiTac': doi_tac,
                    'MaGD': ma_gd,
                })
        page.flush_cache()

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# ============================================================================
# 3D. ACB PDF (GIU NGUYEN)
# ============================================================================

ACB_NGAY_CHINH = "hieu_luc"
_ACB_DEFAULT_EDGES = [36.0, 83.4, 139.6, 164.3, 322.7, 401.5, 480.2, 559.0]
_ACB_DATE_RE = re.compile(r'\d{2}/\d{2}/\d{4}')
_ACB_NUM_RE = r'(-?[\d.,]+)'


def _acb_amount(s):
    s = (s or "").strip().replace(" ", "")
    if not s:
        return ""
    neg = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.strip("-()")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r'\d{1,3}(,\d{3})+', s) else s.replace(",", ".")
    elif "." in s and re.fullmatch(r'\d{1,3}(\.\d{3})+', s):
        s = s.replace(".", "")
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
        if _acb_is_header_fill(r):
            groups.setdefault(round(r['top']), []).append(r)
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
        if top_edge - t > thr:
            break
        top_edge = t
    bot_edge = anchors[-1]
    for t in sorted(t for t in tops if t > anchors[-1] + 1.0):
        if t - bot_edge > thr:
            break
        bot_edge = t
    edges = [top_edge - 1.0] + cuts + [bot_edge + 1.0]
    return list(zip(edges[:-1], edges[1:]))


def _acb_page_rows(page):
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    edges, table_top = _acb_layout(page, words)

    def col_of(w):
        c = (w['x0'] + w['x1']) / 2
        for i in range(7):
            if edges[i] <= c < edges[i + 1]:
                return i
        return None

    cand = [(w, col_of(w)) for w in words if w['top'] >= table_top]
    cand = [(w, c) for w, c in cand if c is not None]
    anchors = sorted(w['top'] for w, c in cand if c == 2 and re.fullmatch(r'\d+', w['text']))
    if not anchors:
        return []
    bounds = _acb_row_bounds([w['top'] for w, _ in cand], anchors)
    rows = []
    for (lo, hi) in bounds:
        cells = {i: [] for i in range(7)}
        for w, c in cand:
            if lo < w['top'] <= hi:
                cells[c].append(w)

        def txt(i):
            ws = sorted(cells[i], key=lambda w: (round(w['top']), w['x0']))
            return " ".join(w['text'] for w in ws).strip()

        sogd = next((w['text'] for w in cells[2] if re.fullmatch(r'\d+', w['text'])), "")
        m0 = _ACB_DATE_RE.search(txt(0))
        rows.append({
            'hieu_luc': m0.group(0) if m0 else txt(0),
            'giao_dich': txt(1),
            'so_gd': sogd,
            'noi_dung': txt(3),
            'no': _acb_amount(txt(4)),
            'co': _acb_amount(txt(5)),
            'so_du': _acb_amount(txt(6)),
        })
    return rows


def _acb_header(page):
    info = {}
    text = page.extract_text() or ""
    m = re.search(r'Thời gian sao kê\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', text)
    if m:
        info['period'] = f"Tu {m.group(1)} den {m.group(2)}"
    for key, pat in (
        ('beginning', r'Số dư đầu\s+'),
        ('total_debit', r'Tổng tiền rút ra\s+'),
        ('total_credit', r'Tổng tiền gửi vào\s+'),
        ('ending', r'Số dư cuối\s+')
    ):
        mm = re.search(pat + _ACB_NUM_RE, text)
        if mm:
            info[key] = _acb_amount(mm.group(1))
    words = page.extract_words()
    left = [w for w in words if w['x1'] < 145 and 150 < w['top'] < 340]
    line_tops = sorted({round(w['top'], 1) for w in left})
    labels = {}
    for t in line_tops:
        s = " ".join(w['text'] for w in sorted((w for w in left if abs(w['top'] - t) < 1.5), key=lambda w: w['x0']))
        key = ('name' if s.startswith('Chủ') else 'acc' if s.startswith('Số tài khoản') else None)
        if key:
            labels.setdefault(key, []).append(t)
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
                'STT': str(stt),
                'Ngay': normalize_date_ddmmyyyy(ngay_chinh),
                'MoTa': r['noi_dung'],
                'No': r['no'] or "0",
                'Co': r['co'] or "0",
                'SoDu': r['so_du'],
                'DoiTac': "",
                'MaGD': r['so_gd'],
            })
        page.flush_cache()
    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': hdr.get('beginning'), 'ending_balance': hdr.get('ending'),
        'total_debit': hdr.get('total_debit'), 'total_credit': hdr.get('total_credit'),
    }


# ============================================================================
# 3E. MBBANK PDF
# ============================================================================

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
                    'STT': row[0],
                    'Ngay': normalize_date_ddmmyyyy(row[1] or row[2]),
                    'MoTa': row[7],
                    'No': row[3] or "0",
                    'Co': row[4] or "0",
                    'SoDu': row[5],
                    'DoiTac': row[6],
                    'MaGD': row[8],
                })
        page.flush_cache()
    all_rows.reverse()
    for i, r in enumerate(all_rows, start=1):
        r['STT'] = str(i)
    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# ============================================================================
# 3F. VIETCOMBANK PDF (SO PHU GIAY BAO NO/CO + TACH PHI + SAO KE MAU 1 & 2)
# ============================================================================

def _normalize_doc_no(s):
    if not s:
        return ""
    s = str(s).strip()
    m_full = re.search(r'(\d{1,6})[.\-\s]+(\d{3,14})$', s)
    if m_full:
        return f"{m_full.group(1)}-{m_full.group(2)}"
    return re.sub(r'\s+', '', s)


def parse_vcb_sophu_advices(pdf):
    """
    Doc tung trang Giay bao No / Giay bao Co trong So phu Vietcombank & Giay bao no VCB Digibiz.
    Dong thoi boc tach chinh xac khoan PHI CHUYEN TIEN (gồm VAT) bi gop tren Sao ke!
    """
    advices = []
    owner_name_top = ""
    owner_acc_top = ""

    for page in pdf.pages:
        text = unicodedata.normalize("NFC", page.extract_text() or "")
        if "SỔ HẠCH TOÁN CHI TIẾT" in text or "DETAILED ACCOUNT BOOK" in text:
            m_acc = re.search(r'Tài khoản\s*/\s*Account\s*:\s*(\S+)', text)
            if m_acc:
                owner_acc_top = m_acc.group(1).strip()
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            for i, ln in enumerate(lines):
                if "CIF" in ln and i + 1 < len(lines) and "Địa chỉ" not in lines[i + 1]:
                    cand_name = lines[i + 1]
                    if i + 2 < len(lines) and "Địa chỉ" not in lines[i + 2] and len(lines[i + 2]) < 15:
                        cand_name += " " + lines[i + 2]
                    owner_name_top = cand_name.strip()
                    break
            page.flush_cache()
            continue

        is_debit = ("GIẤY BÁO NỢ" in text or "DEBIT ADVICE" in text)
        is_credit = ("GIẤY BÁO CÓ" in text or "CREDIT ADVICE" in text)
        if not (is_debit or is_credit):
            page.flush_cache()
            continue

        direction = "out" if is_debit else "in"

        m_date = re.search(r'(?:Ngày\s*/\s*Date|Ngày giao dịch|Transaction date|Ngày hiệu lực|Value date)\s*[:/]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{4})', text, re.I)
        date_val = normalize_date_ddmmyyyy(m_date.group(1)) if m_date else ""

        m_doc = re.search(r'(?:Số chứng từ\s*/\s*Doc No|Số tham chiếu\s*/\s*Reference(?:\s*No)?|Số giao dịch)\s*[:.]?\s*([0-9]+\s*[.\-]\s*[0-9]+|[A-Z0-9.\-]{5,30})', text, re.I)
        doc_key = _normalize_doc_no(m_doc.group(1)) if m_doc else ""

        m_det = re.search(r'(?:Nội dung\s*/\s*Details|Nội dung thanh toán|Diễn giải)\s*:\s*(.*)', text, re.I)
        details = m_det.group(1).strip() if m_det else ""

        ord_name = ord_acc = ord_mst = ""
        ben_name = ben_acc = ben_mst = ""
        amt_debit_str = amt_credit_str = ""
        explicit_fee = 0
        explicit_vat = 0
        explicit_principal = 0
        explicit_total = 0

        def _extract_val(cell_text):
            if ":" in cell_text:
                return cell_text.split(":", 1)[1].strip()
            return ""

        for table in page.extract_tables():
            for row in table:
                cells = [re.sub(r'\s+', ' ', (c or "").replace("\n", " ")).strip() for c in row]
                if len(cells) < 2:
                    continue
                mid = len(cells) // 2
                left_str = " ".join(cells[:mid]).strip()
                right_str = " ".join(cells[mid:]).strip()

                if "Họ tên" in left_str or "A/C Name" in left_str:
                    ord_name = _extract_val(left_str)
                elif "Số TK" in left_str or "A/C No" in left_str:
                    ord_acc = _extract_val(left_str)
                elif "MST" in left_str or "ID No" in left_str:
                    ord_mst = _extract_val(left_str)
                elif "Số tiền nợ" in left_str or "Debit Amount" in left_str:
                    amt_debit_str = _extract_val(left_str)

                if "Họ tên" in right_str or "A/C Name" in right_str:
                    ben_name = _extract_val(right_str)
                elif "Số TK" in right_str or "A/C No" in right_str:
                    ben_acc = _extract_val(right_str)
                elif "MST" in right_str or "ID No" in right_str:
                    ben_mst = _extract_val(right_str)
                elif "Số tiền có" in right_str or "Credit Amount" in right_str:
                    amt_credit_str = _extract_val(right_str)

                # Quet them dong Phi / VAT trong bang (neu co)
                full_row_str = " | ".join(cells)
                low_row = strip_accents(full_row_str).lower()
                if any(k in low_row for k in ("so tien phi", "phi chuyen tien", "phi giao dich", "tong phi", "fee + vat", "phi / charge", "charge")):
                    m_f = re.findall(r'\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{2})?|\b\d{3,9}\b', full_row_str)
                    if m_f:
                        explicit_fee = max(explicit_fee, int(round(to_num(_vcb_comma_dec_to_std_str(m_f[-1])))))
                if ("vat" in low_row or "thue gtgt" in low_row) and "fee + vat" not in low_row:
                    m_v = re.findall(r'\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{2})?|\b\d{2,8}\b', full_row_str)
                    if m_v:
                        v_cand = int(round(to_num(_vcb_comma_dec_to_std_str(m_v[-1]))))
                        if v_cand <= 500000:
                            explicit_vat = max(explicit_vat, v_cand)

        # Quet them text toan trang cho mau Giay bao no VCB Digibiz / iB@nking (khong phai bang 2 cot So phu)
        t_lines = [l.strip() for l in text.split("\n") if l.strip()]
        for idx_l, ln in enumerate(t_lines):
            low_l = strip_accents(ln).lower()
            # Bang ngang: So tien giao dich/ Amount | Tong phi + thue/ Fee + VAT | Tong so tien/ Total amount
            if ("so tien" in low_l or "amount" in low_l) and ("phi" in low_l or "fee" in low_l) and ("tong" in low_l or "total" in low_l) and idx_l + 1 < len(t_lines):
                nums_next = [int(round(to_num(_vcb_comma_dec_to_std_str(x)))) for x in re.findall(r'\d[\d.,]*', t_lines[idx_l + 1])]
                nums_next = [x for x in nums_next if x >= 0]
                if len(nums_next) >= 3 and nums_next[0] + nums_next[1] == nums_next[2]:
                    explicit_principal, explicit_fee, explicit_vat, explicit_total = nums_next[0], nums_next[1], 0, nums_next[2]
                elif len(nums_next) >= 4 and nums_next[0] + nums_next[1] + nums_next[2] == nums_next[3]:
                    explicit_principal, explicit_fee, explicit_vat, explicit_total = nums_next[0], nums_next[1], nums_next[2], nums_next[3]

            if not ben_name and any(k in low_l for k in ("don vi huong", "don vi thu huong", "ten nguoi huong", "nguoi thu huong", "beneficiary")):
                if ":" in ln:
                    ben_name = ln.split(":", 1)[1].strip()
            if not ben_acc and any(k in low_l for k in ("tai khoan huong", "tai khoan nhan", "so tai khoan nguoi huong", "credit account")):
                m_ba = re.search(r'[:\s]([0-9A-Z]{6,22})\b', ln)
                if m_ba:
                    ben_acc = m_ba.group(1).strip()
            if not ord_acc and any(k in low_l for k in ("tai khoan trich no", "tai khoan ghi no", "tai khoan nguon", "debit account")):
                m_oa = re.search(r'[:\s]([0-9]{8,16})\b', ln)
                if m_oa:
                    ord_acc = m_oa.group(1).strip()

            nums_ln = [int(round(to_num(_vcb_comma_dec_to_std_str(x)))) for x in re.findall(r'\
