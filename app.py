# -*- coding: utf-8 -*-
"""
TRICH XUAT SAO KE NGAN HANG (PDF / HTML / EXCEL / CSV) -> EXCEL
- PDF: VietinBank, BIDV, LPBank (7 cot & 10 cot), ACB (A Chau - fix cat dong & cot chuan xac)
- HTML: VPBank (T24 Core)
- Excel / CSV: Sacombank (fix loi ngay 17-07-2025), Agribank IPCAS, Excel khac
- Dinh dang ngay chuan: dd/mm/yyyy
- Web Flask: Xem truoc chi tiet tung file & tim kiem nhanh truoc khi tai Excel.
"""
import base64
import csv
from html.parser import HTMLParser
import os
import re
import tempfile
import webbrowser
from threading import Timer

from flask import Flask, jsonify, render_template_string, request
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
import pdfplumber
import xlrd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


# =========================================================================
#  1. CHUAN HOA NGAY (dd/mm/yyyy) & XU LY TIEN TE AN TOAN
# =========================================================================
_ENG_MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
}

def normalize_date_ddmmyyyy(val):
    if not val:
        return ""
    s = str(val).strip()

    # 1. Dang tieng Anh: 18 MAR 26 hoac 18 MAR 2026 (VPBank)
    m_eng = re.search(r'(?<!\d)(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2,4})(?!\d)', s)
    if m_eng:
        mon = _ENG_MONTHS.get(m_eng.group(2).upper())
        if mon:
            d, y = int(m_eng.group(1)), int(m_eng.group(3))
            y = y + 2000 if y < 100 else y
            return f"{d:02d}/{mon:02d}/{y:04d}"

    # 2. Uu tien yyyy-mm-dd hoac yyyy/mm/dd truoc (fix VietinBank)
    m_iso = re.search(r'(?<!\d)(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?!\d)', s)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{m:02d}/{y:04d}"

    # 3. Dang dd/mm/yyyy hoac dd-mm-yyyy
    m_num4 = re.search(r'(?<!\d)(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?!\d)', s)
    if m_num4:
        d, m, y = int(m_num4.group(1)), int(m_num4.group(2)), int(m_num4.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{m:02d}/{y:04d}"

    # 4. Dang dd/mm/yy hoac dd-mm-yy
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
    """Kiem tra chuoi co thuc su la so tien hop le khong (chan ngay thang nhu 17-07-2025)"""
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
#  2. NHAN DIEN NGAN HANG TU PDF
# =========================================================================
def detect_bank(text):
    t = text or ""
    if "Ngân Hàng TMCP Công Thương Việt Nam" in t or "eFAST" in t or "vietinbank.vn" in t:
        return "vietinbank"
    if "NGÂN HÀNG TMCP ĐẦU TƯ VÀ PHÁT TRIỂN VIỆT NAM" in t or "Bank for Investment and Development" in t:
        return "bidv"
    if "LPBank" in t or "Loc Phat Viet Nam" in t or "Lộc Phát Việt Nam" in t:
        return "lpbank"
    if "www.acb.com.vn" in t or "NGÂN HÀNG THƯƠNG MẠI CỔ PHẦN Á CHÂU" in t or "ASIA COMMERCIAL BANK" in t:
        return "acb"
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
            if "Chi nhánh Branch:" in line: header_info['branch'] = line.split("Branch:", 1)[1].strip()
            if line.startswith("Ngân Hàng"): header_info['bank'] = line.strip()
            if "Từ ngày" in line and "Đến ngày" in line: header_info['period'] = line.strip()
            if "Dear Customer:" in line: header_info['customer'] = line.split("Dear Customer:", 1)[1].strip()
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
                'STT': anchor['text'],
                'Ngay': ngay_chuan,
                'MoTa': txt("MoTa"),
                'No': txt("No"),
                'Co': txt("Co"),
                'SoDu': txt("SoDu"),
                'ThongTinThem': (txt("DoiTac") + " " + txt("SoGD")).strip(),
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
        if line.startswith("Chi nhánh/Branch:"): header_info['branch'] = line.split(":", 1)[1].strip()
        elif line.startswith("NGÂN HÀNG"): header_info['bank'] = line
        elif "Từ ngày/From Date:" in line: header_info['period'] = line
        elif line.startswith("Khách hàng/"):
            m = re.search(r'Customer:\s*(.*?)\s+Địa chỉ', line)
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
                    if val: beginning_balance = val
                    continue
                if "Cộng phát sinh" in label or "Total Amount" in label:
                    nums = [c for c in row[1:] if c not in (None, "")]
                    if len(nums) >= 2: total_debit, total_credit = nums[0], nums[1]
                    continue
                if "Số dư cuối kỳ" in label or "Closing balance" in label:
                    nums = [c for c in row[1:] if c not in (None, "")]
                    if nums: ending_balance = nums[-1] if len(nums) == 1 else nums[0]
                    continue
                if len(row) < 11: continue
                stt = row[0]
                if not stt or not stt.isdigit(): continue

                def clean(v): return (v or "").replace("\n", "").strip()
                extra = " | ".join(filter(None, [
                    f"Mã GD: {clean(row[3])}" if clean(row[3]) else "",
                    f"Chứng từ: {clean(row[7])}" if clean(row[7]) else "",
                    f"GDV: {clean(row[8])}" if clean(row[8]) else "",
                    f"CN: {clean(row[9])}" if clean(row[9]) else "",
                ]))

                all_rows.append({
                    'STT': stt,
                    'Ngay': normalize_date_ddmmyyyy((row[1] or "").replace("\n", " ").strip()),
                    'MoTa': (row[10] or "").replace("\n", " ").strip(),
                    'No': row[4] or "0",
                    'Co': row[5] or "0",
                    'SoDu': row[6] or "",
                    'ThongTinThem': extra,
                })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3C. LPBANK PDF (7 cot moi & 10 cot cu + ghep dong qua trang)
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
        if "Sao kê chi tiết từ" in line:
            header_info['period'] = line
        elif line.startswith("Tên tài khoản:"):
            m = re.search(r'Tên tài khoản:\s*(.*?)\s+(?:Mã khách hàng|Account Name):\s*(\S+)?', line)
            if m:
                header_info['account_name'] = m.group(1).strip()
                if m.group(2): header_info['cif_no'] = m.group(2).strip()
            else:
                header_info['account_name'] = line.split(":", 1)[1].strip()
        elif line.startswith("Số tài khoản:"):
            m = re.search(r'Số tài khoản:\s*(\S+)', line)
            if m: header_info['account_no'] = m.group(1)
        elif line.startswith("Tên khách hàng:"):
            header_info['customer'] = line.split(":", 1)[1].strip()
        elif line.startswith("Đơn vị mở tài khoản:"):
            m = re.search(r'Đơn vị mở tài khoản:\s*(.*?)\s+Đơn vị in:', line)
            header_info['branch'] = m.group(1).strip() if m else line.split(":", 1)[1].strip()
    header_info['bank'] = "LPBank (Ngan hang TMCP Loc Phat Viet Nam)"

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
                    thong_tin_them = " | ".join(filter(None, [
                        f"Hiệu lực: {row[1]}" if row[1] and row[1] != row[0] else "",
                        f"TK doi ung: {row[3]}" if row[3] else "",
                        f"Doi tac: {row[4]}" if row[4] else "",
                        f"NH doi ung: {row[5]}" if row[5] else "",
                        f"So GD: {row[6]}" if row[6] else "",
                    ]))
                    no_val, co_val, sodu_val = row[7] or "0", row[8] or "0", row[9]
                else:
                    thong_tin_them = " | ".join(filter(None, [
                        f"Hiệu lực: {row[1]}" if row[1] and row[1] != row[0] else "",
                        f"So GD: {row[3]}" if row[3] else "",
                    ]))
                    no_val, co_val, sodu_val = row[4] or "0", row[5] or "0", row[6]

                all_rows.append({
                    'STT': str(stt_counter),
                    'Ngay': normalize_date_ddmmyyyy(row[0]),
                    'MoTa': row[2],
                    'No': no_val,
                    'Co': co_val,
                    'SoDu': sodu_val,
                    'ThongTinThem': thong_tin_them,
                })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3D. ACB PDF (DA FIX: Tu dong tim dung tieu de bang + dung dai nen xam cat dong)
# =========================================================================
_ACB_MONEY_RE = re.compile(r'^\d{1,3}(\.\d{3})*$')
_ACB_DATE_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')

def extract_acb(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    p0_text = pdf.pages[0].extract_text() or ""
    for line in p0_text.split("\n"):
        line = line.strip()
        if line.startswith("Chủ tài khoản"): header_info['customer'] = line.replace("Chủ tài khoản", "", 1).strip()
        elif line.startswith("Số tài khoản"):
            m = re.search(r'(\d{5,})', line)
            if m: header_info['account_no'] = m.group(1)
        elif line.startswith("Thời gian sao kê"): header_info['period'] = line.replace("Thời gian sao kê", "", 1).strip()
        elif line.startswith("Số dư đầu"):
            m = re.search(r'([\d.]+)$', line)
            if m: beginning_balance = _vn_money_to_std_str(m.group(1))
        elif line.startswith("Tổng tiền rút ra"):
            m = re.search(r'([\d.]+)$', line)
            if m: total_debit = _vn_money_to_std_str(m.group(1))
        elif line.startswith("Tổng tiền gửi vào"):
            m = re.search(r'([\d.]+)$', line)
            if m: total_credit = _vn_money_to_std_str(m.group(1))
        elif line.startswith("Số dư cuối"):
            m = re.search(r'([\d.]+)$', line)
            if m: ending_balance = _vn_money_to_std_str(m.group(1))
    header_info['bank'] = "ACB (Ngan hang TMCP A Chau)"

    p0_words = pdf.pages[0].extract_words(use_text_flow=False, keep_blank_chars=False)

    # Tim dong tieu de bang giao dich (phai co chu "Nội" va chu "rút" / "gửi" gan cung hang ngang, tranh hop tom tat)
    header_y = None
    for w in p0_words:
        if w['text'] == 'Nội' and 120 < w['x0'] < 220:
            if any(w2['text'] in ('rút', 'gửi', 'tiền') and w2['x0'] > 300 and abs(w2['top'] - w['top']) < 18 for w2 in p0_words):
                header_y = w['top']
                break

    if header_y is None:
        header_y = 320.0

    hdr_words = [w for w in p0_words if (header_y - 12) <= w['top'] <= (header_y + 28)]
    p0_table_top = max([w['bottom'] for w in hdr_words], default=header_y + 25) + 2

    # Tinh ranh gioi cot chinh xac tu tieu de
    col1_words = [w for w in hdr_words if w['x0'] < 75]
    col2_words = [w for w in hdr_words if 75 <= w['x0'] < 128]
    col3_words = [w for w in hdr_words if 128 <= w['x0'] < 158]
    col4_words = [w for w in hdr_words if 158 <= w['x0'] < 250]

    x_hl_end = ((max([w['x1'] for w in col1_words], default=72) + min([w['x0'] for w in col2_words], default=80)) / 2)
    x_gd_end = ((max([w['x1'] for w in col2_words], default=120) + min([w['x0'] for w in col3_words], default=134)) / 2)
    x_sogd_end = ((max([w['x1'] for w in col3_words], default=150) + min([w['x0'] for w in col4_words], default=160)) / 2)

    # Tim toa do mep phai 3 cot tien tu tieu de
    x1_rut = max([w['x1'] for w in hdr_words if w['text'] in ('ra', 'rút', 'tiền') and 340 < w['x0'] < 430], default=405.0)
    x1_gui = max([w['x1'] for w in hdr_words if w['text'] in ('vào', 'gửi', 'tiền') and 430 < w['x0'] < 500], default=480.0)
    x1_du = max([w['x1'] for w in hdr_words if w['text'] in ('dư', 'Số') and w['x0'] > 500], default=557.0)

    mid_rut_gui = (x1_rut + x1_gui) / 2
    mid_gui_du = (x1_gui + x1_du) / 2
    x_money_start = x1_rut - 75

    stt_counter = 0
    for pi, page in enumerate(pdf.pages):
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        table_top = p0_table_top if pi == 0 else 15.0

        footer_tops = [w['top'] for w in words if (w['text'] == 'NGÂN' and w['x0'] > 300) or (w['text'] == '/' and w['top'] > 760)]
        table_bottom = min(footer_tops) - 5 if footer_tops else page.height - 25

        twords = [w for w in words if table_top <= w['top'] < table_bottom]

        # Tim neo ngay o cot 1 (Ngay hieu luc) va loc trung lap theo chieu doc
        raw_anchors = sorted([w for w in twords if w['x0'] < x_hl_end and _ACB_DATE_RE.match(w['text'])], key=lambda w: (w['top'], w['x0']))
        anchors = []
        for a in raw_anchors:
            if not anchors or abs(a['top'] - anchors[-1]['top']) >= 10:
                anchors.append(a)
        if not anchors:
            continue

        # Thu doc cac dai nen mau xam (rects/curves) de lay ranh gioi tung dong chinh xac tuyet doi
        raw_shapes = [s for s in (list(page.rects) + list(page.curves))
                      if s['width'] > 400 and 14 < s['height'] < 250 and s['top'] >= table_top - 4 and s['bottom'] <= table_bottom + 4]
        bg_shapes = []
        for s in sorted(raw_shapes, key=lambda x: x['top']):
            if not bg_shapes or abs(s['top'] - bg_shapes[-1]['top']) > 4:
                bg_shapes.append(s)

        row_spans = []
        if len(bg_shapes) >= max(1, len(anchors) // 3):
            for i, a in enumerate(anchors):
                ay = (a['top'] + a['bottom']) / 2
                inside = next((s for s in bg_shapes if (s['top'] - 2) <= ay <= (s['bottom'] + 2)), None)
                if inside:
                    y_s, y_e = inside['top'] - 1, inside['bottom'] + 1
                else:
                    above = [s for s in bg_shapes if s['bottom'] <= ay + 2]
                    below = [s for s in bg_shapes if s['top'] >= ay - 2]
                    y_s = (above[-1]['bottom'] - 1) if above else table_top
                    y_e = (below[0]['top'] + 1) if below else table_bottom
                row_spans.append((y_s, y_e))
        else:
            # Du phong neu PDF khong co nen xam: cat theo khe ho lon nhat giua cac dong mo ta
            cuts = [table_top]
            for i in range(len(anchors) - 1):
                y_a, y_b = anchors[i]['top'], anchors[i + 1]['top']
                desc_w = [w for w in twords if (y_a - 2) <= w['top'] <= (y_b + 2) and x_sogd_end <= w['x0'] < x_money_start]
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
            row_spans = [(cuts[i], cuts[i + 1]) for i in range(len(anchors))]

        for i in range(len(anchors)):
            y_start, y_end = row_spans[i]
            row_words = [w for w in twords if y_start <= w['top'] < y_end]
            hl_parts, gd_parts, sogd_parts, mota_parts = [], [], [], []
            no_val, co_val, sodu_val = "0", "0", ""

            for w in sorted(row_words, key=lambda x: (round(x['top']), x['x0'])):
                txt = w['text']
                if w['x0'] < x_hl_end:
                    hl_parts.append(txt)
                elif w['x0'] < x_gd_end:
                    gd_parts.append(txt)
                elif w['x0'] < x_sogd_end:
                    sogd_parts.append(txt)
                elif w['x1'] > x_money_start and _ACB_MONEY_RE.match(txt):
                    if w['x1'] < mid_rut_gui: no_val = _vn_money_to_std_str(txt)
                    elif w['x1'] < mid_gui_du: co_val = _vn_money_to_std_str(txt)
                    else: sodu_val = _vn_money_to_std_str(txt)
                else:
                    mota_parts.append(txt)

            stt_counter += 1
            ngay_gd_str = " ".join(gd_parts)
            ngay_hl_str = " ".join(hl_parts)
            ngay_chuan = normalize_date_ddmmyyyy(ngay_gd_str or ngay_hl_str)
            gio_m = re.search(r'\d{2}:\d{2}:\d{2}', ngay_gd_str)

            all_rows.append({
                'STT': str(stt_counter),
                'Ngay': ngay_chuan,
                'MoTa': " ".join(mota_parts).strip(),
                'No': no_val,
                'Co': co_val,
                'SoDu': sodu_val,
                'ThongTinThem': " | ".join(filter(None, [
                    f"Số GD: {' '.join(sogd_parts)}" if sogd_parts else "",
                    f"Giờ: {gio_m.group(0)}" if gio_m else "",
                    f"Hiệu lực: {ngay_hl_str}" if ngay_hl_str and ngay_hl_str != ngay_chuan else "",
                ])),
            })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3E. VPBANK HTML
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
    header_info['bank'] = "VPBank"

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
                'STT': str(stt),
                'Ngay': normalize_date_ddmmyyyy(val_date),
                'MoTa': get(row, "NARR") or pdesc,
                'No': (get(row, "PDAMT") or "0").strip().lstrip("-").strip() or "0",
                'Co': (get(row, "PCAMT") or "0").strip().lstrip("-").strip() or "0",
                'SoDu': get(row, "BOOKED BALANCE"),
                'ThongTinThem': f"Loại GD: {pdesc} | Số tham chiếu: {get(row, 'REFNO')}",
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
#  3F. EXCEL / CSV (SACOMBANK DA FIX LOI 17-07-2025 + AGRIBANK + CHUNG)
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
        except Exception:
            continue
    return []


def extract_excel_or_csv(path):
    rows = _read_tabular_light(path)
    if not rows:
        raise ValueError("File Excel/CSV rong hoac khong doc duoc")

    # 1. AGRIBANK IPCAS
    header_idx = -1
    for i, r in enumerate(rows[:6]):
        lower = [str(c).replace("\ufeff", "").strip().lower() for c in r]
        if "trdt" in lower and ("acctccyamt" in lower or "trccyamt" in lower):
            header_idx = i
            break

    if header_idx != -1:
        col = {str(c).replace("\ufeff", "").strip().lower(): idx for idx, c in enumerate(rows[header_idx])}
        header_info, all_rows, stt = {'bank': 'Agribank (IPCAS)'}, [], 0
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
                'No': f"{abs(amt):,.0f}" if amt < 0 else "0",
                'Co': f"{amt:,.0f}" if amt > 0 else "0",
                'SoDu': f"{to_num(bal):,.0f}" if bal else "",
                'ThongTinThem': f"Tham chieu: {r[col.get('thrref', 0)]} | So GD: {r[col.get('trno', 0)]}",
            })
        if all_rows:
            return {'bank': 'agribank', 'header_info': header_info, 'rows': all_rows,
                    'beginning_balance': f"{to_num(beg_bal):,.0f}" if beg_bal else None,
                    'ending_balance': f"{to_num(end_bal):,.0f}" if end_bal else None,
                    'total_debit': None, 'total_credit': None}

    # 2. SACOMBANK (Da fix bo qua dong thong tin dau trang co ngay mo tai khoan 17-07-2025)
    vnd_rows = [r for r in rows if "VND" in [str(c).strip() for c in r]]
    if len(vnd_rows) >= 2:
        header_info, all_rows, stt = {'bank': 'Sacombank'}, [], 0
        for r in rows[:25]:
            line = " ".join(str(c) for c in r)
            m = re.search(r'(?:Số|Tài khoản):\s*([0-9A-Za-z]+)', line)
            if m and not header_info.get('account_no'):
                header_info['account_no'] = m.group(1)

        for r in vnd_rows:
            r_str = [str(c).strip() for c in r]
            v_idx = r_str.index("VND")
            desc = r_str[v_idx - 1] if v_idx > 0 else ""

            # Bo qua dong metadata "Loại tiền: VND Ngày mở: 17-07-2025"
            if "Loại tiền" in desc or "Loại tiền" in " ".join(r_str[:v_idx]):
                continue

            # Bat buoc phai co ngay giao dich o truoc cot VND moi xu ly so tien
            d_val = ""
            for v in r_str[:v_idx]:
                if re.search(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', v):
                    d_val = normalize_date_ddmmyyyy(v)
                    break
            if not d_val:
                continue

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
                all_rows.append({
                    'STT': str(stt), 'Ngay': d_val, 'MoTa': desc,
                    'No': no_str, 'Co': co_str, 'SoDu': "",
                    'ThongTinThem': f"Ma GD: {r_str[2] if len(r_str) > 2 else ''}"
                })
        if all_rows:
            return {'bank': 'sacombank', 'header_info': header_info, 'rows': all_rows, 'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None}

    # 3. EXCEL / CSV CHUNG
    all_rows, stt = [], 0
    for r in rows:
        r_str = [str(x).strip() for x in r]
        d_val = next((normalize_date_ddmmyyyy(x) for x in r_str if re.match(r'^(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})', x)), None)
        if not d_val: continue
        nums = [_vn_money_to_std_str(x) for x in r_str if _is_money_token(x) and _vn_money_to_std_str(x) != "0"]
        texts = [x for x in r_str if len(x) > 5 and not _is_money_token(x) and not re.match(r'^(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})', x)]
        if nums or texts:
            stt += 1
            all_rows.append({'STT': str(stt), 'Ngay': d_val, 'MoTa': texts[0] if texts else "", 'No': nums[0] if nums else "0", 'Co': nums[1] if len(nums) > 1 else "0", 'SoDu': "", 'ThongTinThem': ""})

    return {'bank': 'excel_khac', 'header_info': {'bank': 'Excel/CSV'}, 'rows': all_rows, 'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None}


# =========================================================================
#  4. DIEU PHOI, DOI CHIEU & XUAT EXCEL
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
    while name.lower() in used or name.lower() == "tong quan":
        n += 1
        suffix = f" ({n})"
        name = base[:31 - len(suffix)] + suffix
    used.add(name.lower())
    return name


def build_xlsx(file_paths, out_path, password=None):
    FONT = "Arial"
    header_font = Font(name=FONT, bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    normal_font = Font(name=FONT, size=10)
    bold_font = Font(name=FONT, size=10, bold=True)
    link_font = Font(name=FONT, size=10, color="0563C1", underline="single")
    err_font = Font(name=FONT, size=10, bold=True, color="C0392B")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_fmt = '#,##0'
    ten_bank_map = {
        "vietinbank": "VietinBank", "bidv": "BIDV", "lpbank": "LPBank", "acb": "ACB",
        "vpbank": "VPBank", "agribank": "Agribank", "sacombank": "Sacombank", "excel_khac": "Excel/CSV"
    }

    all_results = []
    for p in file_paths:
        res = extract_any(p, password)
        warns = [] if res.get('loi') else validate(res)
        all_results.append((res, warns))

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Tong quan"
    used_names = {"tong quan"}
    sheet_names = []
    headers = ["STT", "Ngay giao dich", "Mo ta / Dien giai", "No (Debit)", "Co (Credit)", "So du (Balance)", "Thong tin them"]
    widths = [6, 16, 45, 15, 15, 16, 45]
    HEADER_ROW = 6

    for res, warns in all_results:
        name = make_sheet_name(res['source_file'], used_names)
        sheet_names.append(name)
        ws = wb.create_sheet(name)
        hi = res.get('header_info', {})
        ten_bank = ten_bank_map.get(res['bank'], res['bank'])

        tieu_de = res['source_file'] if res.get('loi') else f"{ten_bank} - {hi.get('customer') or hi.get('account_name') or ''}".strip(" -")
        ws.cell(1, 1, tieu_de).font = Font(name=FONT, bold=True, size=13)
        ws.cell(2, 1, f"File: {res['source_file']}   |   So TK: {hi.get('account_no', '')}   |   {hi.get('period', '')}").font = normal_font

        if res.get('loi'):
            ws.cell(3, 1, f"LOI: {res['loi']}").font = err_font
            continue
        elif warns:
            ws.cell(3, 1, "CANH BAO DOI CHIEU: " + " | ".join(warns)).font = err_font
        else:
            ws.cell(3, 1, f"Doi chieu: khop - {len(res['rows'])} giao dich").font = Font(name=FONT, size=10, color="2E7D32")

        for c, h in enumerate(headers, start=1):
            cell = ws.cell(HEADER_ROW, c, h)
            cell.font, cell.fill, cell.border = header_font, header_fill, border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        r = HEADER_ROW + 1
        if res.get('beginning_balance'):
            ws.cell(r, 3, "So du dau ky").font = bold_font
            ws.cell(r, 6, to_num(res['beginning_balance'])).number_format = num_fmt
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).font = bold_font
                ws.cell(r, c).border = border
            r += 1

        for gd in res['rows']:
            sodu_val = to_num(gd['SoDu']) if gd.get('SoDu') else ""
            vals = [int(gd['STT']) if str(gd['STT']).isdigit() else gd['STT'], gd['Ngay'], gd['MoTa'],
                    to_num(gd['No']), to_num(gd['Co']), sodu_val, gd['ThongTinThem']]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(r, c, v)
                cell.font, cell.border = normal_font, border
                cell.alignment = Alignment(vertical="top", wrap_text=(c in (3, 7)))
            for c in (4, 5, 6):
                if ws.cell(r, c).value != "": ws.cell(r, c).number_format = num_fmt
            r += 1

        if res.get('ending_balance'):
            ws.cell(r, 3, "So du cuoi ky").font = bold_font
            ws.cell(r, 6, to_num(res['ending_balance'])).number_format = num_fmt
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).font = bold_font
                ws.cell(r, c).border = border
            r += 1

        if res.get('total_debit') or res.get('total_credit'):
            ws.cell(r, 3, "Cong phat sinh (theo sao ke)")
            ws.cell(r, 4, to_num(res.get('total_debit'))).number_format = num_fmt
            ws.cell(r, 5, to_num(res.get('total_credit'))).number_format = num_fmt
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).font = bold_font
                ws.cell(r, c).border = border
            r += 1

        for i, w in enumerate(widths, start=1): ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = ws.cell(HEADER_ROW + 1, 1)

    ws_sum.cell(1, 1, "TONG QUAN CAC SAO KE").font = Font(name=FONT, bold=True, size=14)
    sum_headers = ["Sheet (bam de mo)", "File nguon", "Ngan hang", "Khach hang", "So tai khoan", "Ky sao ke",
                   "So du dau ky", "So du cuoi ky", "Tong No", "Tong Co", "So giao dich", "Doi chieu"]
    sum_widths = [24, 34, 14, 34, 16, 30, 16, 16, 18, 18, 12, 40]
    for c, h in enumerate(sum_headers, start=1):
        cell = ws_sum.cell(3, c, h)
        cell.font, cell.fill, cell.border = header_font, header_fill, border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, (res, warns) in enumerate(all_results):
        rr = 4 + i
        hi = res.get('header_info', {})
        status = ("LOI: " + res['loi']) if res.get('loi') else (("CANH BAO: " + " | ".join(warns[:3])) if warns else "Khop")
        t_deb = to_num(res['total_debit']) if res.get('total_debit') else sum(to_num(x['No']) for x in res['rows'])
        t_cred = to_num(res['total_credit']) if res.get('total_credit') else sum(to_num(x['Co']) for x in res['rows'])

        vals = [sheet_names[i], res['source_file'], ten_bank_map.get(res['bank'], res['bank']),
                hi.get('customer') or hi.get('account_name') or '', hi.get('account_no', ''), hi.get('period', ''),
                to_num(res['beginning_balance']) if res.get('beginning_balance') else None,
                to_num(res['ending_balance']) if res.get('ending_balance') else None,
                t_deb, t_cred, len(res['rows']), status]

        for c, v in enumerate(vals, start=1):
            cell = ws_sum.cell(rr, c, v)
            cell.font, cell.border = normal_font, border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

        link_cell = ws_sum.cell(rr, 1)
        link_cell.hyperlink = Hyperlink(ref=link_cell.coordinate, location=f"'{sheet_names[i]}'!A1", display=sheet_names[i])
        link_cell.font = link_font
        for c in (7, 8, 9, 10): ws_sum.cell(rr, c).number_format = num_fmt
        if res.get('loi') or warns: ws_sum.cell(rr, 12).font = err_font

    for i, w in enumerate(sum_widths, start=1): ws_sum.column_dimensions[get_column_letter(i)].width = w
    ws_sum.freeze_panes = "A4"

    wb.save(out_path)
    return all_results


# =========================================================================
#  5. GIAO DIEN WEB FLASK (XEM TRUOC CHI TIET + TIM KIEM)
# =========================================================================
HTML_UI = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Trích Xuất & Xem Trước Sao Kê Đa Ngân Hàng</title>
    <style>
        * { box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, sans-serif; margin: 0; padding: 20px; background: #f0f2f5; color: #333; }
        .container { max-width: 1200px; margin: 0 auto; background: #fff; padding: 28px; border-radius: 12px; box-shadow: 0 4px 18px rgba(0,0,0,0.08); }
        h1 { color: #1F4E78; margin-top: 0; font-size: 23px; }
        .subtitle { color: #555; font-size: 14px; margin-bottom: 18px; line-height: 1.5; }
        .upload-zone { border: 2px dashed #1F4E78; background: #f8fbff; padding: 30px 20px; text-align: center; border-radius: 10px; cursor: pointer; margin-bottom: 14px; }
        .upload-zone:hover { background: #eaf2ff; }
        .controls { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 15px; }
        input[type="password"], input[type="text"] { padding: 10px 12px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px; }
        .btn { padding: 11px 22px; border: none; border-radius: 6px; font-weight: bold; font-size: 14px; cursor: pointer; }
        .btn-primary { background: #1F4E78; color: white; }
        .btn-primary:hover { background: #163a5a; }
        .btn-download { background: #28a745; color: white; text-decoration: none; display: inline-block; }
        .btn-download:hover { background: #218838; }
        #fileList { margin: 8px 0 0 0; font-size: 13px; color: #1F4E78; font-weight: 500; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th, td { border: 1px solid #dee2e6; padding: 8px 10px; text-align: left; }
        th { background: #1F4E78; color: white; position: sticky; top: 0; z-index: 2; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .debit { color: #d9534f; font-weight: 600; }
        .credit { color: #28a745; font-weight: 600; }
        .ok { color: #2e7d32; font-weight: bold; }
        .warn { color: #c0392b; font-weight: bold; }
        .loader { display: none; color: #1F4E78; font-weight: bold; margin: 12px 0; }
        .preview-section { margin-top: 25px; border-top: 2px solid #eee; padding-top: 18px; }
        .tabs { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
        .tab-btn { padding: 8px 14px; border: 1px solid #ccc; background: #f8f9fa; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; }
        .tab-btn.active { background: #1F4E78; color: #fff; border-color: #1F4E78; }
        .preview-toolbar { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 10px; }
        .scroll-box { max-height: 480px; overflow: auto; border: 1px solid #dee2e6; border-radius: 6px; }
    </style>
</head>
<body>
<div class="container">
    <h1>🏦 Công Cụ Trích Xuất & Xem Trước Sao Kê Đa Ngân Hàng</h1>
    <div class="subtitle">
        Hỗ trợ: <b>VietinBank, BIDV, LPBank, ACB (PDF)</b> | <b>VPBank (HTML)</b> | <b>Sacombank, Agribank (Excel/CSV)</b>.<br>
        Chuẩn ngày: <b>dd/mm/yyyy</b>. Kiểm tra trực tiếp từng dòng giao dịch trên Web trước khi tải Excel!
    </div>

    <div class="upload-zone" onclick="document.getElementById('files').click()">
        <input type="file" id="files" multiple accept=".pdf,.html,.htm,.xlsx,.xls,.csv" style="display:none" onchange="showFiles()">
        <div style="font-size: 16px; font-weight: bold; color: #1F4E78;">📂 Bấm vào đây để chọn 1 hoặc nhiều file sao kê</div>
        <div id="fileList"></div>
    </div>

    <div class="controls">
        <input type="password" id="pwd" style="width: 320px; max-width: 100%;" placeholder="🔑 Mật khẩu file PDF (nếu có, không có để trống)">
        <button class="btn btn-primary" onclick="uploadAndProcess()">🚀 BẮT ĐẦU TRÍCH XUẤT & XEM TRƯỚC</button>
    </div>

    <div id="loader" class="loader">⏳ Đang trích xuất và đối chiếu dữ liệu, vui lòng đợi...</div>

    <div id="resultBox" style="display:none;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 12px;">
            <h3 style="margin: 0; color: #1F4E78;">📊 Bảng Đối Chiếu Tổng Quan</h3>
            <a id="downloadLink" class="btn btn-download" href="#" download="KetQua_TongHop_SaoKe.xlsx">📥 TẢI FILE EXCEL TỔNG HỢP (.xlsx)</a>
        </div>

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

        <div class="preview-section">
            <h3 style="margin: 0 0 12px 0; color: #1F4E78;">🔍 Xem Trước Chi Tiết Từng File (Trước Khi Tải)</h3>
            <div id="tabs" class="tabs"></div>
            <div class="preview-toolbar">
                <div id="sheetMeta" style="font-size: 13px; color: #555; font-weight: 500;"></div>
                <input type="text" id="searchBox" style="width: 280px;" placeholder="🔎 Lọc theo ngày, số tiền, nội dung..." oninput="filterPreview()">
            </div>
            <div class="scroll-box">
                <table>
                    <thead>
                        <tr>
                            <th class="text-center" style="width: 50px;">STT</th>
                            <th class="text-center" style="width: 105px;">Ngày</th>
                            <th class="text-right" style="width: 135px;">Nợ (Debit)</th>
                            <th class="text-right" style="width: 135px;">Có (Credit)</th>
                            <th>Mô tả / Diễn giải</th>
                            <th class="text-right" style="width: 135px;">Số dư</th>
                            <th style="width: 200px;">Thông tin thêm</th>
                        </tr>
                    </thead>
                    <tbody id="previewBody"></tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<script>
let allSheetsData = [];
let activeIdx = 0;

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
                <td class="text-right">${r.count}</td><td class="text-right">${r.debit}</td>
                <td class="text-right">${r.credit}</td><td class="${cls}">${r.status}</td>`;
            tbody.appendChild(tr);
        });

        allSheetsData = data.sheets || [];
        renderTabs();

        const dl = document.getElementById('downloadLink');
        dl.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + data.file_b64;
        document.getElementById('resultBox').style.display = 'block';
    } catch (e) {
        alert('Có lỗi xảy ra: ' + e);
    } finally {
        document.getElementById('loader').style.display = 'none';
    }
}

function renderTabs() {
    const tabsEl = document.getElementById('tabs');
    tabsEl.innerHTML = '';
    allSheetsData.forEach((s, idx) => {
        let btn = document.createElement('button');
        btn.className = 'tab-btn' + (idx === 0 ? ' active' : '');
        btn.innerText = `${s.file} (${s.rows.length} dòng)`;
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
        `Ngân hàng: <b>${s.bank}</b> | Số TK: <b>${s.account || '—'}</b> | Tổng số dòng: <b>${s.rows.length}</b>`;
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
        (r.ThongTinThem || '').toLowerCase().includes(kw)
    );
    renderPreviewRows(filtered);
}

function renderPreviewRows(rows) {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';
    if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="color:#888;padding:20px;">Không có dữ liệu giao dịch phù hợp</td></tr>';
        return;
    }
    rows.forEach(r => {
        let tr = document.createElement('tr');
        let noDisp = (r.No && r.No !== '0') ? r.No : '';
        let coDisp = (r.Co && r.Co !== '0') ? r.Co : '';
        tr.innerHTML = `
            <td class="text-center">${r.STT}</td>
            <td class="text-center"><b>${r.Ngay}</b></td>
            <td class="text-right debit">${noDisp}</td>
            <td class="text-right credit">${coDisp}</td>
            <td>${r.MoTa}</td>
            <td class="text-right">${r.SoDu || ''}</td>
            <td style="font-size:12px;color:#555;">${r.ThongTinThem || ''}</td>
        `;
        tbody.appendChild(tr);
    });
}
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML_UI)


@app.route("/process", methods=["POST"])
def process_files():
    uploaded = request.files.getlist("files")
    pwd = request.form.get("password", "").strip() or None
    if not uploaded:
        return jsonify({"error": "Không có file nào được tải lên"}), 400

    with tempfile.TemporaryDirectory() as tmp_dir:
        paths = []
        for f in uploaded:
            p = os.path.join(tmp_dir, os.path.basename(f.filename))
            f.save(p)
            paths.append(p)

        out_xlsx = os.path.join(tmp_dir, "KetQua_TongHop_SaoKe.xlsx")
        results = build_xlsx(paths, out_xlsx, password=pwd)

        with open(out_xlsx, "rb") as f:
            b64_data = base64.b64encode(f.read()).decode("utf-8")

        summary, sheets = [], []
        for res, warns in results:
            hi = res.get('header_info', {})
            t_deb = to_num(res['total_debit']) if res.get('total_debit') else sum(to_num(x['No']) for x in res['rows'])
            t_cred = to_num(res['total_credit']) if res.get('total_credit') else sum(to_num(x['Co']) for x in res['rows'])
            if res.get('loi'): status = "❌ LỖI: " + res['loi']
            elif warns: status = "⚠️ CẢNH BÁO: " + " | ".join(warns[:2])
            else: status = "Khớp chuẩn"

            summary.append({
                "file": res['source_file'],
                "bank": str(res['bank']).upper(),
                "account": hi.get('account_no', ''),
                "count": len(res['rows']),
                "debit": f"{t_deb:,.0f}",
                "credit": f"{t_cred:,.0f}",
                "status": status
            })

            sheets.append({
                "file": res['source_file'],
                "bank": str(res['bank']).upper(),
                "account": hi.get('account_no', ''),
                "rows": res['rows']
            })

    return jsonify({"summary": summary, "sheets": sheets, "file_b64": b64_data})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    if not os.environ.get("RENDER") and not os.environ.get("PYTHONANYWHERE_DOMAIN"):
        Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host="0.0.0.0", port=port)
