# -*- coding: utf-8 -*-
"""
CONG CU TRICH XUAT SAO KE NGAN HANG
- PDF: VietinBank, BIDV, LPBank (7 cot & 10 cot), ACB, MBBank,
       Vietcombank (VCB - iB@nking co ke o, Sao ke khong ke o, va So phu Giay bao No/Co)
- HTML: VPBank (T24 Core)
- Excel/CSV: Sacombank, Agribank IPCAS, Excel khac
- Tu dong tach Phi chuyen tien (Fee + VAT) tu Giay bao No Vietcombank thanh dong rieng (6422 - vietcombank)
- 3 mau ket xuat Excel:
  1. Mau 4 cot (Ngay, No, Co, Dien giai)
  2. Ket xuat cho ASP (6 cot: + TK doi ung, Ma doi tuong)
  3. Ket xuat MISA SME (1 file gom 2 sheet: Phieu_thu_tien_gui & Phieu_chi_tien_gui)
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


def _vcb_comma_dec_to_std_str(val):
    s = str(val or "").strip()
    if re.match(r'^-?\d{1,3}(\.\d{3})+(,\d+)?$', s):
        s = s.replace(".", "").replace(",", ".")
        try:
            num = abs(float(s))
            return f"{num:,.0f}" if num != 0 else "0"
        except ValueError:
            return "0"
    return _vn_money_to_std_str(s)


# =========================================================================
#  2. NHAN DIEN NGAN HANG TU NOI DUNG FILE PDF
# =========================================================================
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
    if "GIẤY BÁO NỢ - DEBIT ADVICE" in t or "GIẤY BÁO CÓ - CREDIT ADVICE" in t or "SỔ HẠCH TOÁN CHI TIẾT" in t:
        return "vcb_sophu"
    if "www.vietcombank.com.vn" in t or "VIETCOMBANK" in t.upper():
        if "VCB - iB@nking" in t or "Nội dung chi tiết" in t or "Transaction In Detail" in t:
            return "vietcombank_vcb2"
        if "Số tham chiếu" in t or "Số tiền ghi nợ" in t:
            return "vietcombank_vcb1"
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
#  3F. VIETCOMBANK PDF (SO PHU + TACH PHI + SAO KE MAU 1 & MAU 2)
# =========================================================================
def _normalize_doc_no(s):
    if not s:
        return ""
    s = str(s).strip()
    m_full = re.search(r'(\d{4})[.\-\s]+(\d{4,12})$', s)
    if m_full:
        return f"{m_full.group(1)}-{m_full.group(2)}"
    return re.sub(r'\s+', '', s)


def parse_vcb_sophu_advices(pdf):
    """Doc tung trang Giay bao No / Giay bao Co trong So phu Vietcombank (bao gom ca tach Phi + VAT)"""
    advices = []
    owner_name_top = ""
    owner_acc_top = ""

    for page in pdf.pages:
        text = unicodedata.normalize("NFC", page.extract_text() or "")

        if "SỔ HẠCH TOÁN CHI TIẾT" in text or "DETAILED ACCOUNT BOOK" in text:
            m_acc = re.search(r'Tài khoản\s*/\s*Account\s*:\s*(\S+)', text)
            if m_acc: owner_acc_top = m_acc.group(1).strip()
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            for i, ln in enumerate(lines):
                if "CIF" in ln and i + 1 < len(lines) and "Địa chỉ" not in lines[i + 1]:
                    cand_name = lines[i + 1]
                    if i + 2 < len(lines) and "Địa chỉ" not in lines[i + 2] and len(lines[i + 2]) < 15:
                        cand_name += " " + lines[i + 2]
                    owner_name_top = cand_name.strip()
                    break
            continue

        is_debit = ("GIẤY BÁO NỢ" in text or "DEBIT ADVICE" in text)
        is_credit = ("GIẤY BÁO CÓ" in text or "CREDIT ADVICE" in text)
        if not (is_debit or is_credit):
            continue

        direction = "out" if is_debit else "in"
        m_date = re.search(r'Ngày\s*/\s*Date\s*:\s*(\d{2}/\d{2}/\d{4})', text)
        date_val = m_date.group(1) if m_date else ""

        m_doc = re.search(r'Số chứng từ\s*/\s*Doc No\s*([0-9.]+)', text)
        doc_key = _normalize_doc_no(m_doc.group(1)) if m_doc else ""

        m_det = re.search(r'Nội dung\s*/\s*Details\s*:\s*(.*)', text)
        details = m_det.group(1).strip() if m_det else ""

        ord_name = ord_acc = ord_mst = ""
        ben_name = ben_acc = ben_mst = ""
        amt_debit_str = amt_credit_str = fee_str = vat_str = ""

        def _extract_val(cell_text):
            return cell_text.split(":", 1)[1].strip() if ":" in cell_text else ""

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
                    m = re.search(r'(?:Số tiền nợ|Debit Amount)\s*:\s*([\d.,]+)', left_str)
                    if m: amt_debit_str = m.group(1)
                elif "Số tiền phí" in left_str or "Fee" in left_str:
                    m = re.search(r'(?:Số tiền phí|Fee)\s*:\s*([\d.,]+)', left_str)
                    if m: fee_str = m.group(1)
                elif "Trong đó VAT" in left_str or "VAT" in left_str:
                    m = re.search(r'(?:Trong đó VAT|VAT)\s*:\s*([\d.,]+)', left_str)
                    if m: vat_str = m.group(1)

                if "Họ tên" in right_str or "A/C Name" in right_str:
                    ben_name = _extract_val(right_str)
                elif "Số TK" in right_str or "A/C No" in right_str:
                    ben_acc = _extract_val(right_str)
                elif "MST" in right_str or "ID No" in right_str:
                    ben_mst = _extract_val(right_str)
                elif "Số tiền có" in right_str or "Credit Amount" in right_str:
                    m = re.search(r'(?:Số tiền có|Credit Amount)\s*:\s*([\d.,]+)', right_str)
                    if m: amt_credit_str = m.group(1)

        owner_acc = (ord_acc if direction == "out" else ben_acc) or owner_acc_top
        owner_name = (ord_name if direction == "out" else ben_name) or owner_name_top
        partner_name = ben_name if direction == "out" else ord_name
        partner_acc = ben_acc if direction == "out" else ord_acc
        partner_mst = ben_mst if direction == "out" else ord_mst

        amt_debit = int(round(to_num(_vcb_comma_dec_to_std_str(amt_debit_str))))
        amt_credit = int(round(to_num(_vcb_comma_dec_to_std_str(amt_credit_str))))
        fee_val = int(round(to_num(_vcb_comma_dec_to_std_str(fee_str))))
        vat_val = int(round(to_num(_vcb_comma_dec_to_std_str(vat_str))))
        fee_total = fee_val + vat_val

        amt_main = amt_debit if direction == "out" else amt_credit
        amt_alt = amt_credit if direction == "out" else amt_debit

        advices.append({
            "owner_acc": owner_acc.strip(),
            "owner_name": owner_name.strip(),
            "date": date_val,
            "doc_key": doc_key,
            "direction": direction,
            "amount": amt_main,
            "alt_amount": amt_alt,
            "fee_total": fee_total,
            "partner_name": partner_name.strip(),
            "partner_acc": partner_acc.strip(),
            "partner_mst": partner_mst.strip(),
            "details": details
        })
    return advices


def extract_vietcombank_vcb1(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    p0_text = pdf.pages[0].extract_text() or ""
    lines0 = [l.strip() for l in p0_text.split("\n") if l.strip()]
    for i, line in enumerate(lines0):
        if line.startswith("Chủ tài"):
            m = re.search(r'khoản:\s*(.+)', line)
            if m: header_info['customer'] = m.group(1).strip()
            elif i + 1 < len(lines0) and "khoản:" in lines0[i + 1]:
                header_info['customer'] = lines0[i + 1].split("khoản:", 1)[1].strip()
        if "Số tài" in line and i + 1 < len(lines0) and "khoản:" in lines0[i + 1]:
            m = re.search(r'(\d{6,})', lines0[i + 1])
            if m: header_info['account_no'] = m.group(1)
        elif re.match(r'^Số tài khoản:\s*(\S+)', line):
            header_info['account_no'] = line.split(":", 1)[1].strip()

    if not header_info.get('account_no'):
        m_acc = re.search(r'\b(\d{9,14})\b', p0_text)
        if m_acc: header_info['account_no'] = m_acc.group(1)
    if not header_info.get('customer'):
        for ln in lines0:
            if ln.startswith("CT ") or ln.startswith("CONG TY") or ln.startswith("CÔNG TY"):
                header_info['customer'] = ln.strip()
                break

    p0 = pdf.pages[0]
    p0_words = p0.extract_words(use_text_flow=False, keep_blank_chars=False)

    for w in p0_words:
        if w['text'] == 'đầu' and w['top'] < 320:
            vals = sorted([x for x in p0_words if abs(x['top'] - w['top']) < 18 and 95 < x['x0'] < 180], key=lambda x: (round(x['top']), x['x0']))
            if vals: beginning_balance = "".join(x['text'] for x in vals)
        elif w['text'] == 'cuối' and w['top'] < 320:
            vals = sorted([x for x in p0_words if abs(x['top'] - w['top']) < 20 and 245 < x['x0'] < 330], key=lambda x: (round(x['top']), x['x0']))
            if vals: ending_balance = "".join(x['text'] for x in vals)

    hdr_y = next((w['top'] for w in p0_words if w['text'] == 'tham' and any(w2['text'] == 'chiếu' and abs(w2['top'] - w['top']) < 4 for w2 in p0_words)), 260.0)
    hdr_band = [w for w in p0_words if (hdr_y - 14) <= w['top'] <= (hdr_y + 18)]
    p0_table_top = max([w['bottom'] for w in hdr_band], default=hdr_y + 16) + 2

    x_date_end = next((w['x0'] - 4 for w in hdr_band if w['text'] == 'tham'), 115.0) - 8
    x_ref_end = next((w['x1'] + 4 for w in hdr_band if w['text'] == 'chiếu'), 190.0)
    x_co_col_x0 = min([w['x0'] for w in hdr_band if w['text'] in ('có', 'ghi') and w['x0'] > 250], default=262.0)
    x_mota_start = min([w['x0'] for w in hdr_band if w['text'] == 'Mô' and w['x0'] > 300], default=322.0) - 4

    x1_no_hdr = max([w['x1'] for w in hdr_band if w['text'] == 'nợ' and w['x0'] > 190], default=255.0)
    x_no_end = max(x1_no_hdr + 3, (x1_no_hdr + x_co_col_x0) / 2)

    last_words = pdf.pages[-1].extract_words(use_text_flow=False, keep_blank_chars=False)
    tong_y = next((w['top'] for w in last_words if w['text'] == 'Tổng' and any(w2['text'] == 'số' and abs(w2['top'] - w['top']) < 4 for w2 in last_words)), None)
    if tong_y:
        deb_w = sorted([w for w in last_words if (tong_y - 15) <= w['top'] <= (tong_y + 15) and x_ref_end <= w['x0'] < x_no_end], key=lambda w: (round(w['top']), w['x0']))
        cred_w = sorted([w for w in last_words if (tong_y - 15) <= w['top'] <= (tong_y + 15) and x_no_end <= w['x0'] < x_mota_start], key=lambda w: (round(w['top']), w['x0']))
        if deb_w: total_debit = "".join(w['text'] for w in deb_w)
        if cred_w: total_credit = "".join(w['text'] for w in cred_w)

    for pi, page in enumerate(pdf.pages):
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        table_top = p0_table_top if pi == 0 else 15.0

        footer_tops = [w['top'] for w in words if (w['text'] == 'Tổng' and w['x0'] < 120) or (w['text'] == 'Trân' and w['x0'] < 180)]
        table_bottom = min(footer_tops) - 8 if footer_tops else page.height - 25
        twords = [w for w in words if table_top <= w['top'] < table_bottom]

        anchors = sorted([w for w in twords if w['x0'] < x_date_end and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])], key=lambda w: w['top'])
        if not anchors:
            continue

        cuts = [table_top]
        for i in range(len(anchors) - 1):
            y_a, y_b = anchors[i]['top'], anchors[i + 1]['top']
            desc_w = [w for w in twords if (y_a - 4) <= w['top'] <= (y_b + 4) and w['x0'] >= x_mota_start]
            y_lines = []
            for y in sorted(set(round(w['top'], 1) for w in desc_w)):
                if not y_lines or y - y_lines[-1] > 3.5: y_lines.append(y)
            if len(y_lines) >= 2:
                max_gap, best_cut = -1, (y_a + y_b) / 2
                for k in range(len(y_lines) - 1):
                    if y_lines[k + 1] - y_lines[k] > max_gap:
                        max_gap = y_lines[k + 1] - y_lines[k]
                        best_cut = (y_lines[k] + y_lines[k + 1]) / 2
                cuts.append(best_cut)
            else:
                cuts.append((y_a + y_b) / 2)
        cuts.append(table_bottom)

        for i in range(len(anchors)):
            y_s, y_e = cuts[i], cuts[i + 1]
            rwords = [w for w in twords if y_s <= w['top'] < y_e]
            c_date, c_ref, c_no, c_co, c_desc = [], [], [], [], []
            for w in sorted(rwords, key=lambda x: (round(x['top']), x['x0'])):
                cx = (w['x0'] + w['x1']) / 2
                if cx < x_date_end: c_date.append(w['text'])
                elif cx < x_ref_end: c_ref.append(w['text'])
                elif cx < x_no_end: c_no.append(w['text'])
                elif cx < x_mota_start: c_co.append(w['text'])
                else: c_desc.append(w['text'])

            no_str = "".join(c_no)
            co_str = "".join(c_co)
            all_rows.append({
                'STT': str(len(all_rows) + 1),
                'Ngay': normalize_date_ddmmyyyy(" ".join(c_date)),
                'MoTa': " ".join(c_desc).strip(),
                'No': f"{to_num(no_str):,.0f}" if to_num(no_str) != 0 else "0",
                'Co': f"{to_num(co_str):,.0f}" if to_num(co_str) != 0 else "0",
                'SoDu': "",
                'DoiTac': "",
                'MaGD': _normalize_doc_no(" ".join(c_ref)),
            })

    all_rows.reverse()
    for idx, r in enumerate(all_rows, start=1):
        r['STT'] = str(idx)

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': f"{to_num(beginning_balance):,.0f}" if beginning_balance else None,
        'ending_balance': f"{to_num(ending_balance):,.0f}" if ending_balance else None,
        'total_debit': f"{to_num(total_debit):,.0f}" if total_debit else None,
        'total_credit': f"{to_num(total_credit):,.0f}" if total_credit else None,
    }


def extract_vietcombank_vcb2(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    p0_text = pdf.pages[0].extract_text() or ""
    lines0 = [l.strip() for l in p0_text.split("\n") if l.strip()]
    for i, line in enumerate(lines0):
        if "Tên tài khoản" in line:
            part = line.split(":", 1)[1].split("Ngày thực hiện")[0].strip()
            if not part and i + 1 < len(lines0): part = lines0[i + 1]
            if i + 1 < len(lines0) and "Số tài khoản" not in lines0[i + 1] and "Account" not in lines0[i + 1] and len(lines0[i + 1]) < 20:
                part = (part + " " + lines0[i + 1]).strip()
            header_info['customer'] = part
        elif "Số tài khoản" in line:
            m = re.search(r':\s*(\S+)', line)
            if m: header_info['account_no'] = m.group(1)
        elif "Số dư đầu kỳ" in line:
            m = re.search(r':\s*([\d.]+)', line)
            if m: beginning_balance = _vn_money_to_std_str(m.group(1))

    last_text = pdf.pages[-1].extract_text() or ""
    for line in last_text.split("\n"):
        if line.startswith("Tổng số"):
            nums = re.findall(r'[\d.]{3,}', line)
            if len(nums) >= 2:
                total_debit = _vn_money_to_std_str(nums[0])
                total_credit = _vn_money_to_std_str(nums[1])
        elif line.startswith("Số dư cuối kỳ"):
            m = re.search(r':\s*([\d.]+)', line)
            if m: ending_balance = _vn_money_to_std_str(m.group(1))

    for page in pdf.pages:
        for table in page.extract_tables():
            for row in table:
                cells = [re.sub(r'\s+', ' ', (c or "").replace("\n", " ")).strip() for c in row]
                if len(cells) < 6:
                    continue
                stt = cells[0]
                if not stt or not stt.isdigit():
                    continue
                date_doc = cells[1]
                ngay = normalize_date_ddmmyyyy(date_doc)
                doc_m = re.search(r'(\d+\s*-\s*\d+)', date_doc)
                ma_gd = _normalize_doc_no(doc_m.group(1)) if doc_m else ""

                all_rows.append({
                    'STT': stt,
                    'Ngay': ngay,
                    'MoTa': cells[5],
                    'No': _vn_money_to_std_str(cells[2]),
                    'Co': _vn_money_to_std_str(cells[3]),
                    'SoDu': _vn_money_to_std_str(cells[4]),
                    'DoiTac': "",
                    'MaGD': ma_gd,
                })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3G. VPBANK HTML
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
#  3H. EXCEL / CSV (SACOMBANK + AGRIBANK + CHUNG + DANH SACH KH/NCC)
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
    "tmdv", "tmxd", "xnk", "nk", "xk", "tap", "doan", "ubnd", "cap", "thoat", "nuoc",
    "dtxd", "tvxd", "xdtm", "sxkd"
}

_CANON_REPLACEMENTS = [
    (r'\bcong\s+ty\b', 'cty'),
    (r'\bct\s+(?=tnhh|cp|co\s+phan|tm|dv|dt|xd|sx|dien)', 'cty '),
    (r'\bco\s+phan\b', 'cp'),
    (r'\btrach\s+nhiem\s+huu\s+han\b', 'tnhh'),
    (r'\bthuong\s+mai\b', 'tm'),
    (r'\bdich\s+vu\b', 'dv'),
    (r'\bxay\s+dung\b', 'xd'),
    (r'\bdau\s+tu\b', 'dt'),
    (r'\bsan\s+xuat\b', 'sx'),
    (r'\bvan\s+tai\b', 'vt'),
    (r'\bxuat\s+nhap\s+khau\b', 'xnk'),
    (r'\bnhap\s+khau\b', 'nk'),
    (r'\bxuat\s+khau\b', 'xk'),
    (r'\bchi\s+nhanh\b', 'cn'),
    (r'\btong\s+cong\s+ty\b', 'tct'),
    (r'\bmtv\b', ''),
    (r'\bva\b', ''),
    (r'\btmdv\b', 'tm dv'),
    (r'\bdvtm\b', 'dv tm'),
    (r'\btmxd\b', 'tm xd'),
    (r'\bxdtm\b', 'xd tm'),
    (r'\bdtxd\b', 'dt xd'),
]

def _canon_biz(s):
    norm = strip_accents(s or "").lower()
    norm = re.sub(r'[^a-z0-9\s]', ' ', norm)
    norm = re.sub(r'\s+', ' ', norm).strip()
    for pat, rep in _CANON_REPLACEMENTS:
        norm = re.sub(pat, rep, norm)
    norm = re.sub(r'\bcn\s+(?:tct|tinh|ninh binh|ha nam|hung yen|phu tho|ha noi|viet nam).*$', '', norm).strip()
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

        if len(mst) >= 10 and mst in raw_text:
            return code

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

        if not matched_exact and len(dist_tokens) >= 2:
            hits = [t for t in dist_tokens if t in masked_tokens]
            ratio = len(hits) / len(dist_tokens)
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
    bk = str(bank_key or "").strip().lower()
    if bk.startswith("vietcombank"):
        return "vietcombank"
    known = {"acb", "agribank", "bidv", "vietinbank", "lpbank", "mbbank", "vpbank", "sacombank", "vietcombank"}
    if bk in known:
        return bk
    fn = strip_accents(source_file).lower()
    for b in known:
        if b in fn:
            return b
    return ""

_FULL_BANK_NAME_MAP = {
    "vietcombank": "Ngân hàng TMCP Ngoại thương Việt Nam",
    "vietcombank_vcb1": "Ngân hàng TMCP Ngoại thương Việt Nam",
    "vietcombank_vcb2": "Ngân hàng TMCP Ngoại thương Việt Nam",
    "vietinbank": "Ngân hàng TMCP Công Thương Việt Nam",
    "bidv": "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam",
    "acb": "Ngân hàng TMCP Á Châu",
    "mbbank": "Ngân hàng TMCP Quân đội",
    "lpbank": "Ngân hàng TMCP Lộc Phát Việt Nam",
    "vpbank": "Ngân hàng TMCP Việt Nam Thịnh Vượng",
    "agribank": "Ngân hàng Nông nghiệp và Phát triển Nông thôn Việt Nam",
    "sacombank": "Ngân hàng TMCP Sài Gòn Thương Tín",
}

def _resolve_full_bank_name(bank_key, bank_code=""):
    bk = str(bank_key or "").strip().lower()
    if bk in _FULL_BANK_NAME_MAP:
        return _FULL_BANK_NAME_MAP[bk]
    return _FULL_BANK_NAME_MAP.get(bank_code, bank_code.upper() if bank_code else "")


def classify_for_asp(row, owner_name="", owner_code="", owner_acc="",
                     customers=None, suppliers=None, owners=None, rules=None,
                     bank_code="", acc_to_comp=None):
    no_val = to_num(row.get('No'))
    co_val = to_num(row.get('Co'))
    mota = row.get('MoTa', '')
    doitac = row.get('DoiTac', '')
    doitac_acc = row.get('DoiTacAcc', '')
    doitac_mst = row.get('DoiTacMST', '')
    magd = str(row.get('MaGD', '') or '').strip().upper()

    txt = strip_accents(mota).lower()
    dt_txt = strip_accents(doitac).lower()
    combined = re.sub(r'\s+', ' ', f"{txt} {dt_txt} {doitac_acc} {doitac_mst}").strip()
    canon_combined = _canon_biz(combined)

    # -1. UU TIEN CAO NHAT: QUY LUAT DA LUU CUA RIENG CONG TY
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
        "nop thue", "thue thu nhap", "thue dien tu", "cham nop thue", "kho bac", "kbnn", "kb:1
