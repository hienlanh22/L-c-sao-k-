# -*- coding: utf-8 -*-
"""
TRICH XUAT SAO KE NGAN HANG (PDF / HTML / EXCEL / CSV) -> EXCEL
- Giu nguyen 100% thuat toan pdfplumber & HTMLParser tu TrichXuatSaoKe (3).py:
  VietinBank (PDF), BIDV (PDF), LPBank (PDF), VPBank (HTML).
- Tich hop them Sacombank (Excel/CSV), Agribank IPCAS (ke ca .XLS ruot CSV), Excel khac.
- Chuan hoa dinh dang ngay: dd/mm/yyyy.
"""
import os
import re
import tempfile
from html.parser import HTMLParser
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink


# =========================================================================
#  CHUAN HOA NGAY dd/mm/yyyy & SO TIEN
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
    m_eng = re.search(r'(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2,4})', s)
    if m_eng:
        mon = _ENG_MONTHS.get(m_eng.group(2).upper())
        if mon:
            d = int(m_eng.group(1))
            y = int(m_eng.group(3))
            y = y + 2000 if y < 100 else y
            return f"{d:02d}/{mon:02d}/{y:04d}"

    # 2. Dang so dd/mm/yyyy hoac dd-mm-yyyy (co the kem gio)
    m_num = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})', s)
    if m_num:
        d, m, y = int(m_num.group(1)), int(m_num.group(2)), int(m_num.group(3))
        y = y + 2000 if y < 100 else y
        if 1 <= m <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{m:02d}/{y:04d}"

    # 3. Dang yyyy-mm-dd
    m_iso = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return f"{d:02d}/{m:02d}/{y:04d}"

    return s


def to_num(s):
    if s is None or s == "" or s == "-NIL-":
        return 0.0
    return float(re.sub(r'[^\d.\-]', '', str(s)) or 0)


def _vn_money_to_std_str(val):
    """Doi dinh dang tien VN co dau cham hang nghin (15.400 / 200.000.000) sang 15,400 / 200,000,000"""
    s = str(val or "").strip()
    if not s or s == "-NIL-":
        return "0"
    if re.match(r'^-?\d{1,3}(\.\d{3})+(,\d+)?$', s):
        s = s.replace(".", "").replace(",", ".")
        return f"{abs(float(s)):,.0f}"
    num = to_num(s.replace(",", ""))
    return f"{abs(num):,.0f}" if num != 0 else "0"


# =========================================================================
#  BUOC 1: NHAN DIEN NGAN HANG TU PDF (Giu nguyen 100% code goc)
# =========================================================================
def detect_bank(text):
    t = text or ""
    if "Ngân Hàng TMCP Công Thương Việt Nam" in t or "eFAST" in t or "vietinbank.vn" in t:
        return "vietinbank"
    if "NGÂN HÀNG TMCP ĐẦU TƯ VÀ PHÁT TRIỂN VIỆT NAM" in t or "Bank for Investment and Development" in t:
        return "bidv"
    if "LPBank" in t or "Loc Phat Viet Nam" in t or "Lộc Phát Việt Nam" in t:
        return "lpbank"
    return "unknown"


# =========================================================================
#  BUOC 2A: VIETINBANK PDF
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
    header_info = {}
    all_rows = []
    beginning_balance = ending_balance = total_debit = total_credit = None

    for pi, page in enumerate(pdf.pages):
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        ptext = page.extract_text() or ""

        for line in ptext.split("\n"):
            if "Số dư đầu kỳ" in line:
                m = re.search(r'([\d,]+)\s*$', line)
                if m: beginning_balance = m.group(1)
            if "Chi nhánh Branch:" in line:
                header_info['branch'] = line.split("Branch:", 1)[1].strip()
            if line.startswith("Ngân Hàng"):
                header_info['bank'] = line.strip()
            if "Từ ngày" in line and "Đến ngày" in line:
                header_info['period'] = line.strip()
            if "Dear Customer:" in line:
                header_info['customer'] = line.split("Dear Customer:", 1)[1].strip()
            if "Tài khoản số:" in line:
                m = re.search(r'Tài khoản số:\s*(\S+)', line)
                if m: header_info['account_no'] = m.group(1)
            if "Doanh số phát sinh" in line:
                nums = re.findall(r'[\d,]+', line)
                if len(nums) >= 2:
                    total_debit, total_credit = nums[-2], nums[-1]
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

            all_rows.append({
                'STT': anchor['text'],
                'Ngay': normalize_date_ddmmyyyy(txt("Ngay")),
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
#  BUOC 2B: BIDV PDF
# =========================================================================
def extract_bidv(pdf):
    header_info = {}
    all_rows = []
    beginning_balance = ending_balance = total_debit = total_credit = None

    first_page_text = pdf.pages[0].extract_text() or ""
    for line in first_page_text.split("\n"):
        line = line.strip()
        if line.startswith("Chi nhánh/Branch:"):
            header_info['branch'] = line.split(":", 1)[1].strip()
        elif line.startswith("NGÂN HÀNG"):
            header_info['bank'] = line
        elif "Từ ngày/From Date:" in line:
            header_info['period'] = line
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

                if "STT" in label and "No" in label:
                    continue
                if "Số dư đầu kỳ" in label or "Opening balance" in label:
                    val = next((c for c in row[1:] if c not in (None, "")), None)
                    if val: beginning_balance = val
                    continue
                if "Cộng phát sinh" in label or "Total Amount" in label:
                    nums = [c for c in row[1:] if c not in (None, "")]
                    if len(nums) >= 2:
                        total_debit, total_credit = nums[0], nums[1]
                    continue
                if "Số dư cuối kỳ" in label or "Closing balance" in label:
                    nums = [c for c in row[1:] if c not in (None, "")]
                    if nums: ending_balance = nums[-1] if len(nums) == 1 else nums[0]
                    continue
                if len(row) < 11:
                    continue
                stt = row[0]
                if not stt or not stt.isdigit():
                    continue

                def clean(v):
                    return (v or "").replace("\n", "").strip()

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
#  BUOC 2C: LPBANK PDF
# =========================================================================
def _clean_lpb(v):
    if not isinstance(v, str) or not v:
        return ""
    v = re.sub(r'^[A-Z0-9]{2,8}\n', '', v)
    return v.replace("\n", " ").strip()

_DATE_RE_LPB = re.compile(r'^\d{2}/\d{2}/\d{4}$')

def extract_lpbank(pdf):
    header_info = {}
    all_rows = []
    beginning_balance = ending_balance = total_debit = total_credit = None

    first_page_text = pdf.pages[0].extract_text() or ""
    for line in first_page_text.split("\n"):
        line = line.strip()
        if "Sao kê chi tiết từ" in line:
            header_info['period'] = line
        elif line.startswith("Tên tài khoản:"):
            m = re.search(r'Tên tài khoản:\s*(.*?)\s+Mã khách hàng:\s*(\S+)', line)
            if m:
                header_info['account_name'] = m.group(1).strip()
                header_info['cif_no'] = m.group(2).strip()
        elif line.startswith("Số tài khoản:"):
            m = re.search(r'Số tài khoản:\s*(\S+)', line)
            if m: header_info['account_no'] = m.group(1)
            m2 = re.search(r'Loại tiền:\s*(\S+)', line)
            if m2: header_info['currency'] = m2.group(1)
        elif line.startswith("Tên khách hàng:"):
            header_info['customer'] = line.split(":", 1)[1].strip()
        elif line.startswith("Địa chỉ:"):
            header_info['address'] = line.split(":", 1)[1].strip()
        elif line.startswith("Đơn vị mở tài khoản:"):
            m = re.search(r'Đơn vị mở tài khoản:\s*(.*?)\s+Đơn vị in:', line)
            if m: header_info['branch'] = m.group(1).strip()
    header_info['bank'] = "LPBank (Ngan hang TMCP Loc Phat Viet Nam)"

    stt_counter = 0
    for page in pdf.pages:
        for table in page.extract_tables():
            for row in table:
                row = [_clean_lpb(c) for c in row]
                if len(row) < 10:
                    continue
                col0 = row[0]

                if col0.startswith("Cộng doanh số") or "Total" in col0:
                    total_debit, total_credit = row[7] or None, row[8] or None
                    continue
                if col0.startswith("Số dư cuối") or "Closing Balance" in col0:
                    ending_balance = row[9] or None
                    continue
                if any("Số dư đầu" in c or "Openning Balance" in c for c in row):
                    beginning_balance = row[9] or None
                    continue
                if not _DATE_RE_LPB.match(col0):
                    continue

                stt_counter += 1
                thong_tin_them = " | ".join(filter(None, [
                    f"Hiệu lực: {row[1]}" if row[1] and row[1] != row[0] else "",
                    f"TK doi ung: {row[3]}" if row[3] else "",
                    f"Doi tac: {row[4]}" if row[4] else "",
                    f"NH doi ung: {row[5]}" if row[5] else "",
                    f"So GD: {row[6]}" if row[6] else "",
                ]))
                all_rows.append({
                    'STT': str(stt_counter),
                    'Ngay': normalize_date_ddmmyyyy(row[0]),
                    'MoTa': row[2],
                    'No': row[7] or "0",
                    'Co': row[8] or "0",
                    'SoDu': row[9],
                    'ThongTinThem': thong_tin_them,
                })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  BUOC 2D: VPBANK HTML
# =========================================================================
_VPB_MONTHS = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
               "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}
_VPB_DATE_RE = re.compile(r'^(\d{1,2}) ([A-Za-z]{3}) (\d{2,4})$')


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.title = ""
        self.tables = []
        self._stack = []
        self._row = None
        self._cell = None
        self._in_title = False

    def _close_cell(self):
        if self._cell is not None and self._row is not None:
            txt = re.sub(r'\s+', ' ', "".join(self._cell).replace("\xa0", " ")).strip()
            self._row.append(txt)
        self._cell = None

    def _close_row(self):
        self._close_cell()
        if self._row is not None and self._stack:
            self._stack[-1].append(self._row)
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
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _vpb_date(s):
    m = _VPB_DATE_RE.match((s or "").strip())
    if not m:
        return normalize_date_ddmmyyyy(s)
    mon = _VPB_MONTHS.get(m.group(2).upper())
    if not mon:
        return s
    yy = int(m.group(3))
    yy = yy + 2000 if yy < 100 else yy
    return f"{int(m.group(1)):02d}/{mon:02d}/{yy:04d}"


def _vpb_amount(s):
    return (s or "").strip().lstrip("-").strip()


def detect_bank_html(title, header_text):
    t = (title or "").upper()
    h = (header_text or "").upper()
    if "VPBANK" in h or "VPB" in t:
        return "vpbank"
    return "unknown"


def extract_vpbank_html(parser):
    header_info = {}
    all_rows = []
    beginning_balance = ending_balance = total_debit = total_credit = None

    for table in parser.tables:
        for row in table:
            if not row: continue
            label = row[0].upper()
            if label == "NR." or (len(row) >= 3 and row[0].upper().startswith("STATEMENT")):
                header_info['account_no'] = row[-1]
            elif label == "CUSTOMER":
                header_info['cif_no'] = row[1] if len(row) > 1 else ""
                header_info['customer'] = row[2] if len(row) > 2 else ""
                for i, cell in enumerate(row):
                    if cell.upper() == "CCY:" and i + 1 < len(row):
                        header_info['currency'] = row[i + 1]
                    if cell.upper() == "BOOK:" and i + 1 < len(row):
                        header_info['branch'] = row[i + 1]
            elif label.startswith("BALANCE AT PERIOD START") and len(row) > 1:
                beginning_balance = row[1]
    header_info['bank'] = "VPBank (Ngan hang TMCP Viet Nam Thinh Vuong)"

    data_table = None
    col = {}
    header_idx = None
    for table in parser.tables:
        for i, row in enumerate(table):
            names = [c.strip().upper() for c in row]
            if "VALDESC" in names and "PDESC" in names:
                data_table, header_idx = table, i
                for j, nme in enumerate(names):
                    if nme and nme not in col:
                        col[nme] = j
                break
        if data_table is not None:
            break
    if data_table is None:
        raise ValueError("Khong tim thay bang giao dich (hang tieu de VALDESC/PDESC)")

    def get(row, name):
        j = col.get(name)
        return row[j] if j is not None and j < len(row) else ""

    stt = 0
    for row in data_table[header_idx + 1:]:
        if not row or not any(row):
            continue
        val_date = get(row, "VALDESC")
        pdesc = get(row, "PDESC")
        if _VPB_DATE_RE.match(val_date or ""):
            stt += 1
            post = _vpb_date(get(row, "POST"))
            val = _vpb_date(val_date)
            extra_parts = [
                f"Loai GD: {pdesc}" if pdesc else "",
                f"So tham chieu: {get(row, 'REFNO')}" if get(row, "REFNO") else "",
                f"Ngay hach toan: {post}" if post and post != val else "",
                f"Kenh: {get(row, 'KENH.TT')}" if get(row, "KENH.TT") else "",
                f"So the: {get(row, 'CARD.NO')}" if get(row, "CARD.NO") else "",
                f"ATM ref: {get(row, 'ATM.REF')}" if get(row, "ATM.REF") else "",
                f"TK ao: {get(row, 'VIRTUAL.ACCT.ID')}" if get(row, "VIRTUAL.ACCT.ID") else "",
            ]
            all_rows.append({
                'STT': str(stt),
                'Ngay': val,
                'MoTa': get(row, "NARR"),
                'No': _vpb_amount(get(row, "PDAMT")) or "0",
                'Co': _vpb_amount(get(row, "PCAMT")) or "0",
                'SoDu': get(row, "BOOKED BALANCE"),
                'ThongTinThem': " | ".join(p for p in extra_parts if p),
            })
        elif "TOTAL MOVEMENT" in (pdesc or "").upper():
            total_debit = _vpb_amount(get(row, "PDAMT")) or None
            total_credit = _vpb_amount(get(row, "PCAMT")) or None
        elif any(c.upper().startswith("BALANCE AT PERIOD END") for c in row):
            nums = [c for c in row if re.match(r'^-?[\d,]+(\.\d+)?$', c)]
            if nums:
                ending_balance = nums[0]

    if all_rows:
        header_info['period'] = f"Tu {all_rows[0]['Ngay']} den {all_rows[-1]['Ngay']}"

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


def extract_html(path):
    def _err(bank, msg):
        return {'source_file': os.path.basename(path), 'bank': bank,
                'header_info': {}, 'rows': [], 'beginning_balance': None,
                'ending_balance': None, 'total_debit': None, 'total_credit': None, 'loi': msg}
    try:
        parser = _TableParser()
        parser.feed(_read_text_file(path))
        parser.close()
    except Exception as e:
        return _err('LOI', f"Khong doc duoc file HTML: {e}")

    header_text = " ".join(" ".join(r) for tb in parser.tables[:1] for r in tb)
    bank = detect_bank_html(parser.title, header_text)
    if bank != "vpbank":
        return _err('khong_xac_dinh', "Chua ho tro dinh dang HTML nay (moi ho tro VPBank)")
    try:
        res = extract_vpbank_html(parser)
    except Exception as e:
        return _err(bank, f"Loi khi trich xuat: {e}")
    res['source_file'] = os.path.basename(path)
    res['bank'] = bank
    res['loi'] = None
    return res


# =========================================================================
#  BUOC 2E: EXCEL / CSV (SACOMBANK + AGRIBANK + THONG MINH)
# =========================================================================
def _read_tabular_smart(path):
    with open(path, "rb") as f:
        magic = f.read(8)
    is_zip = magic.startswith(b'PK')
    is_ole = magic.startswith(b'\xD0\xCF\x11\xE0')

    if is_zip:
        return pd.read_excel(path, header=None, dtype=str, engine="openpyxl").fillna("")
    if is_ole:
        return pd.read_excel(path, header=None, dtype=str, engine="xlrd").fillna("")

    for enc in ("utf-8-sig", "utf-8", "cp1258", "latin-1"):
        try:
            return pd.read_csv(path, header=None, dtype=str, encoding=enc).fillna("")
        except Exception:
            continue
    return pd.read_csv(path, header=None, dtype=str, encoding_errors="replace").fillna("")


def extract_excel_or_csv(path):
    df_raw = _read_tabular_smart(path)
    if df_raw is None or df_raw.empty:
        raise ValueError("File Excel/CSV rong hoac khong doc duoc")

    rows = df_raw.values.tolist()

    # ---- 1. AGRIBANK IPCAS ----
    header_idx = -1
    for i, r in enumerate(rows[:6]):
        lower = [str(c).replace("\ufeff", "").strip().lower() for c in r]
        if "trdt" in lower and ("acctccyamt" in lower or "trccyamt" in lower):
            header_idx = i
            break

    if header_idx != -1:
        col = {str(c).replace("\ufeff", "").strip().lower(): idx for idx, c in enumerate(rows[header_idx])}
        header_info = {'bank': 'Agribank (IPCAS)'}
        all_rows, stt = [], 0
        beg_bal = end_bal = None

        for r in rows[header_idx + 1:]:
            if not any(r): continue
            d_val = normalize_date_ddmmyyyy(r[col['trdt']])
            if not re.match(r'^\d{2}/\d{2}/\d{4}$', d_val): continue

            stt += 1
            amt_col = 'acctccyamt' if 'acctccyamt' in col else 'trccyamt'
            amt = to_num(r[col[amt_col]])
            no_str = f"{abs(amt):,.0f}" if amt < 0 else "0"
            co_str = f"{amt:,.0f}" if amt > 0 else "0"

            desc = str(r[col['rem']]).strip() if 'rem' in col and str(r[col['rem']]).strip() else ""
            if not desc and 'trcdnm' in col:
                desc = str(r[col['trcdnm']]).strip()

            bal = str(r[col['aftrbal']]).strip() if 'aftrbal' in col else ""
            if stt == 1 and 'bftrbal' in col:
                beg_bal = str(r[col['bftrbal']]).strip()
            end_bal = bal

            if not header_info.get('account_no') and 'idxactno' in col:
                header_info['account_no'] = str(r[col['idxactno']]).strip()
            if not header_info.get('customer') and 'name_4' in col:
                header_info['customer'] = str(r[col['name_4']]).strip()

            all_rows.append({
                'STT': str(stt),
                'Ngay': d_val,
                'MoTa': desc,
                'No': no_str,
                'Co': co_str,
                'SoDu': f"{to_num(bal):,.0f}" if bal else "",
                'ThongTinThem': f"Tham chieu: {r[col.get('thrref', 0)]} | So GD: {r[col.get('trno', 0)]}",
            })
        if all_rows:
            header_info['period'] = f"Tu {all_rows[0]['Ngay']} den {all_rows[-1]['Ngay']}"
            return {
                'bank': 'agribank', 'header_info': header_info, 'rows': all_rows,
                'beginning_balance': f"{to_num(beg_bal):,.0f}" if beg_bal else None,
                'ending_balance': f"{to_num(end_bal):,.0f}" if end_bal else None,
                'total_debit': None, 'total_credit': None,
            }

    # ---- 2. SACOMBANK ----
    vnd_rows = [r for r in rows if "VND" in [str(c).strip() for c in r]]
    if len(vnd_rows) >= 2:
        header_info = {'bank': 'Sacombank'}
        for r in rows[:25]:
            line = " ".join(str(c) for c in r)
            if "Số:" in line or "Tài khoản:" in line:
                m = re.search(r'(?:Số|Tài khoản):\s*([0-9A-Za-z]+)', line)
                if m: header_info['account_no'] = m.group(1)
            if "Chi nhánh:" in line:
                header_info['branch'] = line.split("Chi nhánh:", 1)[1].split(",")[0].strip()

        all_rows, stt = [], 0
        for r in vnd_rows:
            r_str = [str(c).strip() for c in r]
            v_idx = r_str.index("VND")
            desc = r_str[v_idx - 1] if v_idx > 0 else ""
            d_val = ""
            for val in r_str[:v_idx]:
                if re.search(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', val):
                    d_val = normalize_date_ddmmyyyy(val)
                    break

            no_str, co_str = "0", "0"
            for j in range(v_idx + 1, len(r_str)):
                val = r_str[j]
                cleaned = val.replace(".", "").replace(",", "").replace("-", "").strip()
                if cleaned.isdigit() and int(cleaned) > 0:
                    if j - v_idx <= 4:
                        no_str = _vn_money_to_std_str(val)
                    else:
                        co_str = _vn_money_to_std_str(val)

            if d_val and (no_str != "0" or co_str != "0" or desc):
                stt += 1
                all_rows.append({
                    'STT': str(stt),
                    'Ngay': d_val,
                    'MoTa': desc,
                    'No': no_str,
                    'Co': co_str,
                    'SoDu': "",
                    'ThongTinThem': f"Ma GD: {r_str[2] if len(r_str) > 2 else ''}",
                })
        if all_rows:
            header_info['period'] = f"Tu {all_rows[0]['Ngay']} den {all_rows[-1]['Ngay']}"
            return {
                'bank': 'sacombank', 'header_info': header_info, 'rows': all_rows,
                'beginning_balance': None, 'ending_balance': None,
                'total_debit': None, 'total_credit': None,
            }

    # ---- 3. EXCEL / CSV KHAC ----
    all_rows, stt = [], 0
    for r in rows:
        r_str = [str(x).strip() for x in r]
        d_val = next((normalize_date_ddmmyyyy(x) for x in r_str if re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', x)), None)
        if not d_val: continue

        nums = [_vn_money_to_std_str(x) for x in r_str if re.match(r'^-?[\d,.]+$', x) and to_num(x.replace(".", "")) != 0]
        texts = [x for x in r_str if len(x) > 5 and not re.match(r'^-?[\d,.]+$', x) and not re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', x)]

        if nums or texts:
            stt += 1
            all_rows.append({
                'STT': str(stt),
                'Ngay': d_val,
                'MoTa': texts[0] if texts else "",
                'No': nums[0] if len(nums) > 0 else "0",
                'Co': nums[1] if len(nums) > 1 else "0",
                'SoDu': "",
                'ThongTinThem': "",
            })

    return {
        'bank': 'excel_khac', 'header_info': {'bank': 'Excel/CSV'}, 'rows': all_rows,
        'beginning_balance': None, 'ending_balance': None,
        'total_debit': None, 'total_credit': None,
    }


# =========================================================================
#  BUOC 3: DIEU PHOI & KIEM TRA DOI CHIEU
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

    if ext in (".html", ".htm"):
        return extract_html(path)

    if ext in (".xlsx", ".xls", ".csv"):
        try:
            res = extract_excel_or_csv(path)
            res['source_file'] = source_name
            res['loi'] = None
            return res
        except Exception as e:
            return {
                'source_file': source_name, 'bank': 'LOI', 'header_info': {}, 'rows': [],
                'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None,
                'loi': f"Loi doc Excel/CSV: {e}"
            }

    try:
        pdf = pdfplumber.open(path, password=password or "")
    except Exception as e:
        return {
            'source_file': source_name, 'bank': 'LOI', 'header_info': {}, 'rows': [],
            'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None,
            'loi': f"Khong mo duoc file PDF (co the sai mat khau hoac file hong): {e}",
        }

    with pdf:
        first_text = pdf.pages[0].extract_text() or ""
        last_text = pdf.pages[-1].extract_text() or "" if len(pdf.pages) > 1 else ""
        bank = detect_bank(first_text + "\n" + last_text)
        try:
            if bank == "vietinbank": res = extract_vietinbank(pdf)
            elif bank == "bidv": res = extract_bidv(pdf)
            elif bank == "lpbank": res = extract_lpbank(pdf)
            else:
                return {
                    'source_file': source_name, 'bank': 'khong_xac_dinh', 'header_info': {}, 'rows': [],
                    'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None,
                    'loi': "Khong nhan dien duoc day la sao ke PDF cua ngan hang nao (ho tro: VietinBank, BIDV, LPBank)",
                }
        except Exception as e:
            return {
                'source_file': source_name, 'bank': bank, 'header_info': {}, 'rows': [],
                'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None,
                'loi': f"Loi khi trich xuat PDF: {e}",
            }

    res['source_file'] = source_name
    res['bank'] = bank
    res['loi'] = None
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


# =========================================================================
#  BUOC 4: XUAT EXCEL TONG HOP
# =========================================================================
_BAD_SHEET_CHARS = re.compile(r'[\[\]:*?/\\]')

def make_sheet_name(file_name, used):
    stem = _BAD_SHEET_CHARS.sub(" ", os.path.splitext(file_name)[0]).strip().strip("'")
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
        "vietinbank": "VietinBank", "bidv": "BIDV", "lpbank": "LPBank",
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
    headers = ["STT", "Ngay giao dich", "Mo ta / Dien giai", "No (Debit)",
               "Co (Credit)", "So du (Balance)", "Thong tin them"]
    widths = [6, 16, 45, 15, 15, 16, 45]
    HEADER_ROW = 6

    for res, warns in all_results:
        name = make_sheet_name(res['source_file'], used_names)
        sheet_names.append(name)
        ws = wb.create_sheet(name)
        hi = res.get('header_info', {})
        ten_bank = ten_bank_map.get(res['bank'], res['bank'])

        tieu_de = (res['source_file'] if res.get('loi')
                   else f"{ten_bank} - {hi.get('customer') or hi.get('account_name') or ''}".strip(" -"))
        ws.cell(1, 1, tieu_de).font = Font(name=FONT, bold=True, size=13)
        ws.cell(2, 1, f"File: {res['source_file']}   |   So TK: {hi.get('account_no', '')}   |   {hi.get('period', '')}").font = normal_font

        if res.get('loi'):
            ws.cell(3, 1, f"LOI: {res['loi']}").font = err_font
        elif warns:
            ws.cell(3, 1, "CANH BAO DOI CHIEU: " + " | ".join(warns)).font = err_font
        else:
            ws.cell(3, 1, f"Doi chieu: khop - {len(res['rows'])} giao dich").font = Font(name=FONT, size=10, color="2E7D32")

        if res.get('loi'):
            ws.column_dimensions['A'].width = 60
            continue

        for c, h in enumerate(headers, start=1):
            cell = ws.cell(HEADER_ROW, c, h)
            cell.font, cell.fill, cell.border = header_font, header_fill, border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        r = HEADER_ROW + 1

        def write_special(label, balance_val):
            nonlocal r
            ws.cell(r, 3, label)
            ws.cell(r, 6, to_num(balance_val))
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).font = bold_font
                ws.cell(r, c).border = border
            ws.cell(r, 6).number_format = num_fmt
            r += 1

        if res.get('beginning_balance'):
            write_special("So du dau ky", res['beginning_balance'])

        for gd in res['rows']:
            sodu_val = to_num(gd['SoDu']) if gd.get('SoDu') else ""
            values = [int(gd['STT']) if str(gd['STT']).isdigit() else gd['STT'], gd['Ngay'], gd['MoTa'],
                      to_num(gd['No']), to_num(gd['Co']), sodu_val, gd['ThongTinThem']]
            for c, v in enumerate(values, start=1):
                cell = ws.cell(r, c, v)
                cell.font, cell.border = normal_font, border
                cell.alignment = Alignment(vertical="top", wrap_text=(c in (3, 7)))
            for c in (4, 5, 6):
                if ws.cell(r, c).value != "":
                    ws.cell(r, c).number_format = num_fmt
            r += 1

        if res.get('ending_balance'):
            write_special("So du cuoi ky", res['ending_balance'])

        if res.get('total_debit') or res.get('total_credit'):
            ws.cell(r, 3, "Cong phat sinh (theo sao ke)")
            ws.cell(r, 4, to_num(res.get('total_debit')))
            ws.cell(r, 5, to_num(res.get('total_credit')))
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).font = bold_font
                ws.cell(r, c).border = border
            for c in (4, 5):
                ws.cell(r, c).number_format = num_fmt
            r += 1

        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = ws.cell(HEADER_ROW + 1, 1)

    # Sheet Tong quan
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
        if res.get('loi'): status = "LOI: " + res['loi']
        elif warns: status = "CANH BAO: " + " | ".join(warns[:3])
        else: status = "Khop"

        t_deb = to_num(res['total_debit']) if res.get('total_debit') else sum(to_num(x['No']) for x in res['rows'])
        t_cred = to_num(res['total_credit']) if res.get('total_credit') else sum(to_num(x['Co']) for x in res['rows'])

        vals = [sheet_names[i], res['source_file'], ten_bank_map.get(res['bank'], res['bank']),
                hi.get('customer') or hi.get('account_name') or '', hi.get('account_no', ''),
                hi.get('period', ''),
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
        for c in (7, 8, 9, 10):
            ws_sum.cell(rr, c).number_format = num_fmt
        if res.get('loi') or warns:
            ws_sum.cell(rr, 12).font = err_font

    for i, w in enumerate(sum_widths, start=1):
        ws_sum.column_dimensions[get_column_letter(i)].width = w
    ws_sum.freeze_panes = "A4"

    wb.save(out_path)
    return all_results


# =========================================================================
#  BUOC 5: GIAO DIEN WEB STREAMLIT
# =========================================================================
st.set_page_config(page_title="Trích Xuất Sao Kê Đa Ngân Hàng", page_icon="🏦", layout="wide")

st.title("🏦 CÔNG CỤ TRÍCH XUẤT & TỔNG HỢP SAO KÊ")
st.markdown(
    "Hỗ trợ: **VietinBank, BIDV, LPBank (PDF)** | **VPBank (HTML)** | "
    "**Sacombank, Agribank (Excel/CSV)**. Định dạng ngày chuẩn: **`dd/mm/yyyy`**."
)

uploaded_files = st.file_uploader(
    "📁 Kéo thả hoặc chọn 1 hay nhiều file sao kê cùng lúc",
    type=["pdf", "html", "htm", "xlsx", "xls", "csv"],
    accept_multiple_files=True
)

pdf_password = st.text_input(
    "🔑 Mật khẩu mở file PDF (nếu có, không có thì bỏ trống)",
    type="password"
)

if uploaded_files and st.button("🚀 BẮT ĐẦU TRÍCH XUẤT & TẠO EXCEL", type="primary"):
    with st.spinner("⏳ Đang xử lý và đối chiếu số dư từng giao dịch..."):
        with tempfile.TemporaryDirectory() as tmp_dir:
            temp_paths = []
            for uf in uploaded_files:
                file_path = os.path.join(tmp_dir, os.path.basename(uf.name))
                with open(file_path, "wb") as f:
                    f.write(uf.getbuffer())
                temp_paths.append(file_path)

            out_xlsx = os.path.join(tmp_dir, "KetQua_TongHop_SaoKe.xlsx")
            pwd = pdf_password.strip() if pdf_password else None
            results = build_xlsx(temp_paths, out_xlsx, password=pwd)

            with open(out_xlsx, "rb") as f:
                st.session_state["xlsx_bytes"] = f.read()

            summary_rows = []
            for res, warns in results:
                hi = res.get('header_info', {})
                t_deb = to_num(res['total_debit']) if res.get('total_debit') else sum(to_num(x['No']) for x in res['rows'])
                t_cred = to_num(res['total_credit']) if res.get('total_credit') else sum(to_num(x['Co']) for x in res['rows'])

                if res.get('loi'):
                    status = "❌ LỖI: " + res['loi']
                elif warns:
                    status = "⚠️ CẢNH BÁO: " + " | ".join(warns[:2])
                else:
                    status = "✅ Khớp chuẩn"

                summary_rows.append({
                    "File nguồn": res['source_file'],
                    "Ngân hàng": str(res['bank']).upper(),
                    "Số TK": hi.get('account_no', ''),
                    "Số giao dịch": len(res['rows']),
                    "Tổng Nợ": f"{t_deb:,.0f}",
                    "Tổng Có": f"{t_cred:,.0f}",
                    "Đối chiếu": status
                })
            st.session_state["summary_df"] = pd.DataFrame(summary_rows)
            st.session_state["raw_results"] = results

if "xlsx_bytes" in st.session_state:
    st.success("🎉 Đã trích xuất xong! Bấm nút bên dưới để tải File Excel tổng hợp.")
    st.download_button(
        label="📥 TẢI VỀ FILE EXCEL TỔNG HỢP (.xlsx)",
        data=st.session_state["xlsx_bytes"],
        file_name="KetQua_TongHop_SaoKe.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    st.subheader("📊 Bảng đối chiếu tổng quan")
    st.dataframe(st.session_state["summary_df"], use_container_width=True)

    st.subheader("🔍 Xem trước chi tiết từng file")
    for res, warns in st.session_state.get("raw_results", []):
        if res.get("rows"):
            with st.expander(f"{res['source_file']} ({len(res['rows'])} giao dịch)"):
                df_preview = pd.DataFrame(res["rows"])[["STT", "Ngay", "MoTa", "No", "Co", "SoDu"]]
                st.dataframe(df_preview, use_container_width=True)