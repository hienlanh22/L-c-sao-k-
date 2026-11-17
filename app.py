# -*- coding: utf-8 -*-
"""
CONG CU TRICH XUAT SAO KE NGAN HANG
- PDF: VietinBank, BIDV, LPBank (7 cot & 10 cot), ACB, MBBank
- HTML: VPBank (T24 Core)
- Excel/CSV: Sacombank, Agribank IPCAS, Excel khac
- Dinh dang ngay: dd/mm/yyyy
- Ket xuat 4 cot & Ket xuat cho ASP:
  + 515: Lai tien gui -> MaDT = ten ngan hang
  + 6422: Phi ngan hang / chuyen tien -> MaDT = ten ngan hang
  + 1111: Nop/Rut tien mat, Rut/Chi sec, NGUYEN VAN LAN Chuyen tien -> MaDT trong
  + 131 / 331: Tu dong nhan dien theo danh ba KH/NCC cua tung cong ty + 16 cong ty trong TaiKhoanNH
  + Thu goc/lai/lai phat/phi vay, Thue, BHXH, Noi bo, Luong... -> De trong
- Dong bo danh ba dung chung cho moi may tinh qua GitHub (GITHUB_TOKEN + GITHUB_REPO) hoac Firebase (FIREBASE_DB_URL)
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
            companies[code] = {"code": code, "name": name, "accounts": [], "customers": [], "suppliers": []}
        companies[code]["accounts"].append({"bank": bank, "account": acc})
    if "lanhien" in companies:
        companies["lanhien"]["customers"] = list(_DEFAULT_LANHIEN_CUSTOMERS)
    return {"companies": companies}

_LOCAL_DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "partners_db.json")
_ACC_TO_COMPANY_CODE = {acc.strip(): code for _, code, _, acc in _DEFAULT_ACCOUNTS}

def load_shared_db():
    """Doc danh ba tu Cloud (Firebase / GitHub) hoac file cuc bo, gop voi danh sach 16 cong ty mac dinh"""
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
                if cdata.get("customers"): db["companies"][code]["customers"] = cdata["customers"]
                if cdata.get("suppliers"): db["companies"][code]["suppliers"] = cdata["suppliers"]
                if cdata.get("accounts"): db["companies"][code]["accounts"] = cdata["accounts"]
                if cdata.get("name"): db["companies"][code]["name"] = cdata["name"]
            else:
                db["companies"][code] = cdata
    return db

def save_shared_db(db):
    """Luu danh ba ra file cuc bo + dong bo len Cloud (Firebase / GitHub) de may khac dung chung"""
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
            return True, "Đã lưu đồng bộ lên Firebase Cloud (mọi máy đều dùng chung)"
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
            payload = {"message": "Cap nhat danh ba KH/NCC tu Web", "content": content_b64}
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
            return True, "Đã lưu đồng bộ lên GitHub (mọi máy đều dùng chung)"
        except Exception as e:
            return False, f"Lỗi lưu GitHub: {e}"

    return True, "Đã lưu trên máy chủ Web (lưu ý: nếu Render ngủ sẽ reset về danh sách gốc trừ khi gắn GITHUB_TOKEN / FIREBASE_DB_URL)"


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
#  4. BO NHAN DIEN TAI KHOAN DOI UNG & GHEP MA DOI TUONG CHO ASP
# =========================================================================
_PARTNER_ALIASES = [
    ("loc hoang", "lochoang"), ("lan hien", "lanhien"), ("thinh hien", "thinhhien"),
    ("thuy khanh", "thuykhanh"), ("thanh anh", "thanhanh"), ("minh huy", "minhhuy"),
    ("le trung", "letrung"), ("hung long", "hunglong"), ("thanh phat", "thanhphat"),
    ("khoi huong", "khoihuong"), ("phuc thanh", "phucthanh"), ("hoang anh", "hoanganh"),
    ("ha ken", "haken"), ("hoang nguyen", "hoangnguyen"),
    ("cpcknn ly nhan", "cokhi"), ("co khi nn ly nhan", "cokhi"), ("co khi nong nghiep ly nhan", "cokhi"),
    ("duc long ha anh", "duclonghaanh"), ("minh quang", "minhquang"),
    ("hoa lu", "hoalu"), ("hung phat", "hungphat"), ("thang loan", "thangloan"),
    ("thai hung", "thaihung"), ("trung kien", "trungkien"), ("viet nhat", "vietnhat"),
    ("viet xo", "vietxo"), ("toan hung", "toanhung"), ("hau phat", "hauphat"),
    ("thanh dat", "thanhdat"), ("ngoi sao viet", "ngoisaoviet"), ("thep nam a", "thepnama"),
    ("bim son", "bimson"), ("duc duy", "ducduy"), ("thanh nam", "thanhnam"),
    ("quang na", "quangna"), ("tan hong ha", "tanhongha"), ("xuan hung", "xuanhung"),
    ("sang phuong", "sangphuong"), ("hoa sen home", "hoasenhome"), ("hoa binh", "hoabinh"),
    ("anh kien", "anhkien"), ("ba ky", "baky"), ("duc linh", "duclinh"),
    ("viet y", "viety"), ("huihuang", "huihuang"), ("thien son", "thienson"),
    ("nam duy", "namduy"), ("dong a", "donga"), ("ha anh", "haanh"),
    ("hancom", "hancom"), ("vuong thin", "vuongthin"), ("an phu", "anphu"),
    ("thinh cuong", "thinhcuong"), ("thanh bac", "thanhbac"), ("tue hoa", "tuehoa"),
    ("hung hien", "hunghien"), ("xd 226", "xd226"), ("226 ha nam", "xd226"),
    ("doan tue", "doantue"), ("minh nhat", "minhnhat"), ("3f viet", "3fviet"),
    ("vina agri", "dhvinaagri"), ("mavin", "mavin"), ("sunjin", "sunjin"),
    ("mind group", "mindgroup"), ("tap doan 68", "tapdoan68"), ("manh dat", "manhdat"),
    ("ly ha thanh", "lyhathanh"), ("phu thanh an", "phuthanhan"),
    ("hong ha", "hongha"), ("long truong giang", "longtruonggiang"), ("bao khang", "baokhang"),
    ("tan sang", "tansang"), ("hung thinh", "hungthinh"), ("quan anh", "quananh"),
    ("hoang long", "hoanglong"), ("maresti", "maresti"), ("tacovi", "tacovi"),
    ("phuong trang", "phuongtrang"), ("trung tuyet", "trungtuyet"), ("bui minh", "buiminh"),
    ("mai linh", "mailinh"), ("eco home", "ecohome"), ("hung hai", "hunghai"),
    ("thai son", "thaison"), ("quang huong", "quanghuong"), ("kim thanh", "kimthanh"),
    ("viet phap", "vietphap"), ("nam phat", "namphat"), ("pharaon", "pharaon"),
    ("hoa phat", "hoaphat"), ("sinotruk", "sinotruknhuy"), ("gia bao chau", "giabaochau"),
    ("khanh hung", "khanhhung"), ("hai long", "hailong"), ("huyen linh", "huyenlinh"),
    ("huyen anh", "huyenanh"), ("an binh", "anbinh"), ("ben vung ly nhan", "benvunglynhan"),
    ("minh khoa", "minhkhoa"), ("ngoc bien", "ngocbien"), ("y thuy", "ythuy"),
    ("thang hang", "thanghang"), ("nam thanh", "namthanh"),
    ("kyoto", "kyoto"), ("haka", "haka"), ("phuc chau", "phucchau"),
    ("kiotviet", "kiotviet"), ("tung ngoc", "tungngoc"), ("motell", "motell"),
    ("hoa phat auto", "hoaphatauto"), ("dien luc ly nhan", "dienluclynhan"),
    ("trung tam dang kiem", "ttdangkiem"), ("phat trien quy dat", "ptquydat"),
    ("khang duy", "khangduy"), ("toan gia phat", "toangiaphat"), ("thien hoa an", "thienhoaan"),
    ("atc", "atc"), ("mega", "mega")
]

_STOP_WORDS_NAME = {
    "cong", "ty", "cty", "tnhh", "cp", "co", "phan", "mtv", "tm", "dv", "sx", "xd",
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
    "cong", "ty", "cty", "tnhh", "cp", "co", "phan", "mtv", "ho", "kinh", "doanh", "hkd",
    "trach", "nhiem", "huu", "han", "tu", "van", "dau", "tu", "xay", "dung", "thuong", "mai",
    "dich", "vu", "san", "xuat", "van", "tai", "tm", "dv", "sx", "xd", "va", "dt", "dvtm",
    "tmdv", "tmxd", "xnk", "tap", "doan", "ubnd", "cap", "thoat", "nuoc"
}

_DEFAULT_INTERNAL_PEOPLE = {"nguyen van lan"}

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
    for alias_k, alias_v in _PARTNER_ALIASES:
        if alias_k in clean:
            cands.add(alias_k)
            cands.add(alias_v)
    return cands

def _normalize_partner_search_variants(name):
    norm = strip_accents(name).lower()
    norm = re.sub(r'[^a-z0-9\s]', ' ', norm)
    norm = re.sub(r'\s+', ' ', norm).strip()
    words = norm.split()
    variants = []
    if len(words) >= 2:
        variants.append((norm, len(norm) + 20))
    core_words = []
    for idx, w in enumerate(words):
        if w in _LEGAL_PREFIX_WORDS and (len(words) - idx) > 2:
            continue
        core_words.append(w)
    core = " ".join(core_words).strip()
    if core and len(core) >= 4:
        variants.append((core, len(core) + 10))
    if len(core_words) >= 3:
        tail2 = " ".join(core_words[-2:])
        if len(tail2) >= 5:
            variants.append((tail2, len(tail2)))
        tail3 = " ".join(core_words[-3:])
        if len(tail3) >= 7:
            variants.append((tail3, len(tail3) + 2))
    return variants

def _match_from_directory(search_text_norm, raw_text, partner_list, owner_cands):
    if not partner_list:
        return None
    best_code = None
    best_score = -1

    masked = " " + search_text_norm + " "
    for oc in sorted(owner_cands, key=len, reverse=True):
        if len(oc) >= 4:
            masked = masked.replace(f" {oc} ", " ")

    for item in partner_list:
        code = str(item.get("code", "")).strip()
        name = str(item.get("name", "")).strip()
        mst = re.sub(r'\s+', '', str(item.get("mst", "") or ""))
        if not code or not name:
            continue

        if len(mst) >= 10 and mst in raw_text:
            return code

        for variant, base_score in _normalize_partner_search_variants(name):
            if variant in owner_cands:
                continue
            if re.search(r'(?<![a-z0-9])' + re.escape(variant) + r'(?![a-z0-9])', masked):
                if base_score > best_score:
                    best_score = base_score
                    best_code = code

    return best_code

def _extract_partner_code(mota, doitac, owner_name, owner_code="", owner_acc="",
                          preferred_list=None, secondary_list=None):
    owner_cands = _owner_code_candidates(owner_name, owner_code=owner_code)
    raw_combined = f"{mota} {doitac}"
    norm_combined = re.sub(r'[^a-z0-9\s]', ' ', strip_accents(raw_combined).lower())
    norm_combined = re.sub(r'\s+', ' ', norm_combined).strip()

    # 1. UU TIEN CAO NHAT: Tim trong Danh sach KH / NCC da luu cua cong ty
    if preferred_list:
        matched = _match_from_directory(norm_combined, raw_combined, preferred_list, owner_cands)
        if matched:
            return matched
    if secondary_list:
        matched = _match_from_directory(norm_combined, raw_combined, secondary_list, owner_cands)
        if matched:
            return matched

    # 2. Kiem tra so tai khoan doi ung thuoc danh sach 16 cong ty TaiKhoanNH
    for acc_no, comp_code in _ACC_TO_COMPANY_CODE.items():
        if len(acc_no) >= 7 and acc_no != owner_acc and comp_code not in owner_cands:
            if acc_no in raw_combined:
                return comp_code

    # 3. Tim trong cot DoiTac (VietinBank / LPBank / MBBank)
    doitac_clean = strip_accents(doitac or "").lower()
    if doitac_clean and "tien mat" not in doitac_clean and "tai don vi" not in doitac_clean:
        dt_main = re.split(r'\s*[–\-]\s*a/c:|a/c:', doitac_clean)[0].strip()
        for kw, code in _PARTNER_ALIASES:
            if kw in dt_main and kw not in owner_cands and code not in owner_cands:
                return code
        dt_words = [w for w in re.sub(r'[^a-z0-9\s]', ' ', dt_main).split()
                    if w not in _STOP_WORDS_NAME and not w.isdigit()]
        if dt_words:
            cand_code = "".join(dt_words[:4])
            if cand_code not in owner_cands and len(cand_code) >= 3:
                return cand_code

    # 4. Tim theo tu dien mac dinh trong noi dung MoTa
    desc_clean = strip_accents(mota or "").lower()
    desc_clean = re.sub(r'[_\-;,/]+', ' ', desc_clean)
    desc_masked = desc_clean
    for oc in sorted(owner_cands, key=len, reverse=True):
        if len(oc) >= 4:
            desc_masked = desc_masked.replace(oc, " ")

    for kw, code in _PARTNER_ALIASES:
        if kw in desc_masked and kw not in owner_cands and code not in owner_cands:
            return code

    # 5. Tu tach ten cong ty sau "cty" / "cong ty"
    m_cty = re.search(r'(?:cong ty|cty)\s+(?:tnhh|cp|co phan|tm|dv|sx|xd|tmdv|xnk|\s)*([a-z0-9\s]{3,35})', desc_masked)
    if m_cty:
        words = [w for w in m_cty.group(1).split() if w not in _STOP_WORDS_NAME and not re.match(r'^\d{4,}$', w)]
        if words:
            code = "".join(words[:3])
            if code not in owner_cands and len(code) >= 3:
                return code

    return ""


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
                     customers=None, suppliers=None, bank_code=""):
    no_val = to_num(row.get('No'))
    co_val = to_num(row.get('Co'))
    mota = row.get('MoTa', '')
    doitac = row.get('DoiTac', '')
    magd = str(row.get('MaGD', '') or '').strip().upper()

    txt = strip_accents(mota).lower()
    dt_txt = strip_accents(doitac).lower()
    combined = f"{txt} {dt_txt}"

    # 0. CAC KHOAN VAY / THUE / BHXH / NOI BO / LUONG -> DE TRONG ""
    # Chay TRUOC bo loc phi 6422 de "Thu goc: ... Thu lai: ... Thu phi: VND" khong bi bat nham vao 6422!
    other_non_keywords = (
        "thu goc", "thu lai", "thu lai phat", "khe uoc", "tra von vay", "tra lai vay",
        "loan repayment", "thau chi", "dong goc lai", "giai ngan",
        "nop thue", "thue thu nhap", "thue dien tu", "cham nop thue", "kho bac", "kbnn", "kb:1327",
        "bhxh", "bao hiem", "chuyen noi bo", "chi luong",
        "ve duong", "vetc", "chuyen tra lai tien", "nop nham", "giao dich nham"
    )
    if any(k in combined for k in other_non_keywords):
        return "", ""

    # 1. LAI TIEN GUI (TK 515) -> MaDT = ten ngan hang (acb, agribank, bidv...)
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

    # 3. NOP / RUT TIEN MAT, RUT / CHI SEC, NGUYEN VAN LAN Chuyen tien (TK 1111) -> MaDT de trong
    is_internal_person = any(
        re.search(r'(?<![a-z0-9])' + re.escape(pname) + r'(?![a-z0-9])', combined)
        for pname in _DEFAULT_INTERNAL_PEOPLE
    )
    is_cash_or_cheque = (
        "tien mat, vang tien te tai don vi" in dt_txt
        or is_internal_person
        or any(k in combined for k in (
            "nop tien", "rut tien", "rut sec", "chi sec", "linh tien",
            "nt vao tk", "nop tm", "nt-tk-acb", "nop tien mat"
        ))
        or txt.strip() in ("nt", "rt")
        or (bool(re.search(r'\b(rt|nt)\b', txt)) and not any(k in txt for k in ("cty", "cong ty", "tnhh", "hang")))
    )
    if is_cash_or_cheque:
        return "1111", ""

    # 4. KHACH TRA TIEN (TK 131) -> CO > 0
    if co_val > 0 and no_val == 0:
        ma_dt = _extract_partner_code(
            mota, doitac, owner_name, owner_code=owner_code, owner_acc=owner_acc,
            preferred_list=customers, secondary_list=suppliers
        )
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

    # 5. TRA TIEN NHA CUNG CAP (TK 331) -> NO > 0
    if no_val > 0 and co_val == 0:
        ma_dt = _extract_partner_code(
            mota, doitac, owner_name, owner_code=owner_code, owner_acc=owner_acc,
            preferred_list=suppliers, secondary_list=customers
        )
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
#  6. GIAO DIEN WEB PASTEL + DANH BA CHUNG TREN WEB + NUT NOI 3D
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
            max-width: 1260px;
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
        .partner-row { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
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
        .badge-count {
            display: inline-block;
            padding: 5px 11px;
            border-radius: 8px;
            font-size: 12.5px;
            font-weight: 600;
            background: #e8f8f0;
            color: #1f704b;
            border: 1px solid #bfead3;
        }
        .badge-empty { background: #f5f6fa; color: #7c879c; border-color: #e0e4ee; }
        .sync-note { font-size: 12px; color: #6c7a93; margin-top: 8px; }

        .upload-zone {
            border: 2.5px dashed #9cb8f3;
            background: #f4f8ff;
            padding: 28px 20px;
            text-align: center;
            border-radius: 16px;
            cursor: pointer;
            margin-bottom: 18px;
            transition: all 0.2s ease;
        }
        .upload-zone:hover { background: #e9f1ff; border-color: #779be8; transform: translateY(-1px); }
        .controls { display: flex; gap: 14px; align-items: center; flex-wrap: wrap; margin-bottom: 16px; }
        input[type="password"], input[type="text"] {
            padding: 10px 14px;
            border: 1.5px solid #d9e2f2;
            border-radius: 10px;
            font-size: 13.5px;
            background: #fbfcff;
            color: #3d485c;
            outline: none;
        }
        input:focus { border-color: #8eb0f2; }

        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 12px;
            font-weight: 700;
            font-size: 13.5px;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            transition: all 0.15s ease;
            transform: translateY(0);
        }
        .btn-sm { padding: 8px 14px; font-size: 12.5px; border-radius: 10px; }
        .btn-primary {
            background: linear-gradient(180deg, #91b3fa 0%, #769ce9 100%);
            color: #ffffff;
            box-shadow: 0 5px 0 #5b7fc9, 0 8px 16px rgba(118, 156, 233, 0.32);
        }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 7px 0 #5b7fc9, 0 12px 20px rgba(118, 156, 233, 0.38); }
        .btn-primary:active { transform: translateY(3px); box-shadow: 0 2px 0 #5b7fc9, 0 4px 8px rgba(118, 156, 233, 0.25); }

        .btn-download {
            background: linear-gradient(180deg, #83ddb8 0%, #63c79e 100%);
            color: #164f38;
            box-shadow: 0 5px 0 #48a67f, 0 8px 16px rgba(99, 199, 158, 0.32);
        }
        .btn-download:hover { transform: translateY(-2px); box-shadow: 0 7px 0 #48a67f, 0 12px 20px rgba(99, 199, 158, 0.40); }
        .btn-download:active { transform: translateY(3px); box-shadow: 0 2px 0 #48a67f, 0 4px 8px rgba(99, 199, 158, 0.25); }

        .btn-asp {
            background: linear-gradient(180deg, #c2aff7 0%, #a58ee8 100%);
            color: #ffffff;
            box-shadow: 0 5px 0 #8269c9, 0 8px 16px rgba(165, 142, 232, 0.34);
        }
        .btn-asp:hover { transform: translateY(-2px); box-shadow: 0 7px 0 #8269c9, 0 12px 20px rgba(165, 142, 232, 0.42); }
        .btn-asp:active { transform: translateY(3px); box-shadow: 0 2px 0 #8269c9, 0 4px 8px rgba(165, 142, 232, 0.25); }

        .btn-soft {
            background: linear-gradient(180deg, #ffffff 0%, #f1f4fb 100%);
            color: #4b5d7a;
            border: 1px solid #d9e1f2;
            box-shadow: 0 3px 0 #d0d8ea;
        }
        .btn-soft:hover { transform: translateY(-1px); box-shadow: 0 4px 0 #c3cde3; }
        .btn-soft:active { transform: translateY(2px); box-shadow: 0 1px 0 #c3cde3; }

        .btn-peach {
            background: linear-gradient(180deg, #ffd4be 0%, #f8b799 100%);
            color: #693018;
            box-shadow: 0 3px 0 #dc9372;
        }
        .btn-peach:hover { transform: translateY(-1px); box-shadow: 0 4px 0 #dc9372; }
        .btn-peach:active { transform: translateY(2px); box-shadow: 0 1px 0 #dc9372; }

        .tab-btn {
            padding: 9px 16px;
            border: 1px solid #dce3f0;
            background: #f5f7fc;
            color: #4b5d7a;
            border-radius: 10px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 600;
            box-shadow: 0 3px 0 #d5dceb;
            transition: all 0.15s ease;
        }
        .tab-btn:hover { transform: translateY(-1px); background: #ebf0fb; box-shadow: 0 4px 0 #c9d2e6; }
        .tab-btn:active { transform: translateY(2px); box-shadow: 0 1px 0 #c9d2e6; }
        .tab-btn.active {
            background: linear-gradient(180deg, #b8a9f5 0%, #9c8be8 100%);
            color: #ffffff;
            border-color: #8b78db;
            box-shadow: 0 3px 0 #7562c7, 0 5px 10px rgba(156, 139, 232, 0.28);
        }

        .mode-switch {
            display: inline-flex;
            gap: 8px;
            background: #f2f5fc;
            padding: 6px;
            border-radius: 12px;
            border: 1px solid #e1e8f5;
        }
        .mode-pill {
            padding: 7px 14px;
            border-radius: 8px;
            border: none;
            background: transparent;
            color: #5b6982;
            font-weight: 600;
            font-size: 13px;
            cursor: pointer;
        }
        .mode-pill.active {
            background: #ffffff;
            color: #4b659e;
            box-shadow: 0 2px 6px rgba(91, 127, 201, 0.18);
        }

        #fileList { margin: 10px 0 0 0; font-size: 13px; color: #5b7fc9; font-weight: 600; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; background: #fff; }
        th, td { border: 1px solid #e6ebf5; padding: 9px 11px; text-align: left; }
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
        .cell-input {
            width: 100%; padding: 5px 8px; border: 1px solid #dde4f2; border-radius: 6px;
            font-size: 12.5px; background: #fcfdff; color: #3d485c;
        }
        .cell-input:focus { border-color: #9c8be8; background: #fff; outline: none; }
        .tk-badge { text-align: center; font-weight: 700; color: #5b48a8; }
    </style>
</head>
<body>
<div class="container">
    <h1>🌸 Công cụ trích xuất sao kê ngân hàng</h1>
    <div class="subtitle">
        Hỗ trợ: <b>VietinBank, BIDV, LPBank, ACB, MBBank (PDF)</b> • <b>VPBank (HTML)</b> • <b>Sacombank, Agribank (Excel/CSV)</b>.<br>
        Chuẩn ngày: <b>dd/mm/yyyy</b> • 2 mẫu kết xuất: <b>Mẫu 4 cột</b> và <b>Kết xuất cho ASP</b> (<code>1111</code>, <code>6422</code>, <code>131</code>, <code>331</code>, <code>515</code>).
    </div>

    <!-- KHU VUC QUAN LY DANH BA CONG TY DUNG CHUNG TREN WEB -->
    <div class="partner-card">
        <div class="partner-header">
            <div style="font-weight: 700; color: #574699; font-size: 15px;">
                🏢 Danh bạ Công ty, Khách hàng &amp; Nhà cung cấp (Tự động nhận diện theo Số TK sao kê)
            </div>
            <div id="syncBadge" style="font-size: 12px; color: #1f704b; font-weight: 600;"></div>
        </div>

        <div class="partner-row" style="margin-bottom: 12px;">
            <label style="font-weight: 600; font-size: 13px;">Xem / Cập nhật công ty:</label>
            <select id="companySelect" onchange="onSelectCompany()"></select>
            <button class="btn btn-soft btn-sm" onclick="addCompany()">➕ Thêm công ty</button>
            <button class="btn btn-peach btn-sm" onclick="deleteCompany()">🗑️ Xóa công ty</button>
        </div>

        <div class="partner-row">
            <input type="file" id="upKhFile" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadPartnerList('customers', this)">
            <button class="btn btn-soft btn-sm" onclick="triggerPartnerUpload('upKhFile')">📥 Nạp DS Khách hàng (.xlsx)</button>
            <span id="khStatus" class="badge-count badge-empty">0 Khách hàng</span>

            <input type="file" id="upNccFile" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadPartnerList('suppliers', this)">
            <button class="btn btn-soft btn-sm" onclick="triggerPartnerUpload('upNccFile')">📥 Nạp DS Nhà cung cấp (.xlsx)</button>
            <span id="nccStatus" class="badge-count badge-empty">0 Nhà cung cấp</span>
        </div>
        <div class="sync-note" id="syncNoteText">
            ✨ Hệ thống đã nạp sẵn 16 công ty &amp; 29 số tài khoản ngân hàng. Khi tải sao kê lên, web tự dò theo Số tài khoản để ghép đúng danh sách KH/NCC của công ty đó.
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

    <div id="loader" class="loader">⏳ Đang trích xuất, đối chiếu số dư và ghép mã đối tượng...</div>

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
                <h3 style="margin: 0; color: #4b659e;">🔍 Xem Trước &amp; Chỉnh Sửa Trước Khi Tải</h3>
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

            <div class="scroll-box">
                <table>
                    <thead id="previewHead"></thead>
                    <tbody id="previewBody"></tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<script>
let sharedDB = { companies: {} };
let selectedCompanyCode = '';
let allSheetsData = [];
let activeIdx = 0;
let previewMode = 'asp';

async function loadSharedDB() {
    try {
        const res = await fetch('/api/partners');
        const data = await res.json();
        sharedDB = data.db || { companies: {} };
        if (!selectedCompanyCode && sharedDB.companies['lanhien']) {
            selectedCompanyCode = 'lanhien';
        }
        renderCompanyDropdown();
    } catch (e) {
        console.error(e);
    }
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
    updateStatusBadges();
}

function updateStatusBadges() {
    const comp = sharedDB.companies[selectedCompanyCode];
    const khEl = document.getElementById('khStatus');
    const nccEl = document.getElementById('nccStatus');
    if (!comp) {
        khEl.innerText = 'Chọn công ty để xem/nạp KH'; khEl.className = 'badge-count badge-empty';
        nccEl.innerText = 'Chọn công ty để xem/nạp NCC'; nccEl.className = 'badge-count badge-empty';
        return;
    }
    const khCnt = (comp.customers || []).length;
    const nccCnt = (comp.suppliers || []).length;
    khEl.innerText = `${khCnt} Khách hàng`;
    khEl.className = khCnt > 0 ? 'badge-count' : 'badge-count badge-empty';
    nccEl.innerText = `${nccCnt} Nhà cung cấp`;
    nccEl.className = nccCnt > 0 ? 'badge-count' : 'badge-count badge-empty';
}

function onSelectCompany() {
    selectedCompanyCode = document.getElementById('companySelect').value;
    updateStatusBadges();
}

async function persistDBToServer() {
    const res = await fetch('/api/partners/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(sharedDB)
    });
    const data = await res.json();
    document.getElementById('syncBadge').innerText = data.message || '✅ Đã lưu';
}

async function addCompany() {
    const name = prompt('Nhập Tên công ty mới:');
    if (!name || !name.trim()) return;
    const code = prompt('Nhập Mã công ty (viết liền không dấu, VD: abc):');
    if (!code || !code.trim()) return;
    const cleanCode = code.trim().toLowerCase();
    if (!sharedDB.companies[cleanCode]) {
        sharedDB.companies[cleanCode] = { code: cleanCode, name: name.trim(), accounts: [], customers: [], suppliers: [] };
    }
    selectedCompanyCode = cleanCode;
    renderCompanyDropdown();
    await persistDBToServer();
}

async function deleteCompany() {
    if (!selectedCompanyCode) { alert('Hãy chọn một công ty để xóa!'); return; }
    if (!confirm(`Bạn có chắc muốn xóa danh bạ của [${selectedCompanyCode}]?`)) return;
    delete sharedDB.companies[selectedCompanyCode];
    selectedCompanyCode = '';
    renderCompanyDropdown();
    await persistDBToServer();
}

function triggerPartnerUpload(inputId) {
    if (!selectedCompanyCode) {
        alert('Vui lòng chọn Công ty ở ô thả xuống phía trên trước khi nạp danh sách!');
        return;
    }
    document.getElementById(inputId).click();
}

async function uploadPartnerList(roleKey, inputEl) {
    if (!inputEl.files.length || !selectedCompanyCode) return;
    const fd = new FormData();
    fd.append('file', inputEl.files[0]);
    try {
        const res = await fetch('/parse_partners', { method: 'POST', body: fd });
        const data = await res.json();
        if (data.error) { alert(data.error); return; }
        sharedDB.companies[selectedCompanyCode][roleKey] = data.items || [];
        updateStatusBadges();
        await persistDBToServer();
        const label = roleKey === 'customers' ? 'Khách hàng' : 'Nhà cung cấp';
        alert(`✅ Đã nạp & lưu ${data.items.length} ${label} cho [${selectedCompanyCode}] lên hệ thống Web!`);
    } catch (e) {
        alert('Lỗi khi đọc file danh sách: ' + e);
    } finally {
        inputEl.value = '';
    }
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

function selectTab(idx) {
    activeIdx = idx;
    document.querySelectorAll('.tab-btn').forEach((b, i) => b.classList.toggle('active', i === idx));
    document.getElementById('searchBox').value = '';
    const s = allSheetsData[idx];
    document.getElementById('sheetMeta').innerHTML =
        `Ngân hàng: <b>${s.bank}</b> • Số TK: <b>${s.account || '—'}</b> • Công ty: <b>${s.customer || s.company_code || '—'}</b> • Giao dịch: <b>${s.rows.length}</b>`;
    renderPreviewRows(s.rows);
}

function filterPreview() {
    const kw = document.getElementById('searchBox').value.trim().toLowerCase();
    const rows = allSheetsData[activeIdx]?.rows || [];
    if (!kw) { renderPreviewRows(rows); return; }
    const filtered = rows.filter(r =>
        (r.Ngay || '').toLowerCase().includes(kw) ||
        (r.MoTa || '').toLowerCase().includes(kw) ||
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
            <th class="text-center" style="width: 105px;">Ngày</th>
            <th class="text-right" style="width: 130px;">Nợ</th>
            <th class="text-right" style="width: 130px;">Có</th>
            <th>Diễn giải</th>
            <th class="text-center" style="width: 120px;">TK đối ứng</th>
            <th style="width: 155px;">Mã đối tượng</th>
        </tr>`;
    } else {
        thead.innerHTML = `<tr>
            <th class="text-center" style="width: 115px;">Ngày</th>
            <th class="text-right" style="width: 145px;">Nợ</th>
            <th class="text-right" style="width: 145px;">Có</th>
            <th>Diễn giải</th>
        </tr>`;
    }

    const colCount = previewMode === 'asp' ? 6 : 4;
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
                <td>${r.MoTa}</td>
                <td><input class="cell-input tk-badge" value="${r.TK || ''}" onchange="updateRowField(${origIdx}, 'TK', this.value)" placeholder="Trống"></td>
                <td><input class="cell-input" value="${r.MaDT || ''}" onchange="updateRowField(${origIdx}, 'MaDT', this.value)" placeholder="Trống"></td>
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

            # Tu dong nhan dien Ma CT theo So tai khoan trong TaiKhoanNH
            comp_code = _ACC_TO_COMPANY_CODE.get(acc_no, "") or fallback_comp_code
            comp_obj = companies_map.get(comp_code, {})
            owner_name = hi.get('customer') or hi.get('account_name') or comp_obj.get('name', '')
            customers = comp_obj.get("customers", [])
            suppliers = comp_obj.get("suppliers", [])
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
                    customers=customers, suppliers=suppliers, bank_code=bank_code
                )
                preview_rows.append({
                    "Ngay": r['Ngay'],
                    "No": f"{no_v:,.0f}" if no_v != 0 else "0",
                    "Co": f"{co_v:,.0f}" if co_v != 0 else "0",
                    "MoTa": r['MoTa'],
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
