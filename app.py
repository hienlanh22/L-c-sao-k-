# -*- coding: utf-8 -*-
"""
CONG CU TRICH XUAT SAO KE NGAN HANG
- PDF: VietinBank, BIDV, LPBank (7 cot & 10 cot), ACB (fix nhan dien theo ty le trang)
- HTML: VPBank (T24 Core)
- Excel / CSV: Sacombank (fix 17-07-2025), Agribank IPCAS, Excel khac
- Dinh dang ngay: dd/mm/yyyy
- File Excel tai ve chi co 4 cot: Ngay, No, Co, Dien giai
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
import pdfplumber
import xlrd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


# =========================================================================
#  1. CHUAN HOA NGAY (dd/mm/yyyy) & TIEN TE
# =========================================================================
_ENG_MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
}

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
#  2. NHAN DIEN NGAN HANG TU PDF
# =========================================================================
def detect_bank(text):
    t = text or ""
    tl = t.lower()
    if "ngân hàng tmcp công thương việt nam" in tl or "efast" in tl or "vietinbank.vn" in tl:
        return "vietinbank"
    if "ngân hàng tmcp đầu tư và phát triển việt nam" in tl or "bank for investment and development" in tl:
        return "bidv"
    if "lpbank" in tl or "loc phat viet nam" in tl or "lộc phát việt nam" in tl:
        return "lpbank"
    if ("acb.com.vn" in tl or "ngân hàng thương mại cổ phần á châu" in tl
            or "asia commercial bank" in tl or "1900 54 54 86" in tl or "acb signature" in tl):
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
        if line.startswith("Số tài khoản/"):
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

                all_rows.append({
                    'STT': stt,
                    'Ngay': normalize_date_ddmmyyyy((row[1] or "").replace("\n", " ").strip()),
                    'MoTa': (row[10] or "").replace("\n", " ").strip(),
                    'No': row[4] or "0",
                    'Co': row[5] or "0",
                    'SoDu': row[6] or "",
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
        if line.startswith("Số tài khoản:"):
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
                else:
                    no_val, co_val, sodu_val = row[4] or "0", row[5] or "0", row[6]

                all_rows.append({
                    'STT': str(stt_counter),
                    'Ngay': normalize_date_ddmmyyyy(row[0]),
                    'MoTa': row[2],
                    'No': no_val,
                    'Co': co_val,
                    'SoDu': sodu_val,
                })

    return {
        'header_info': header_info, 'rows': all_rows,
        'beginning_balance': beginning_balance, 'ending_balance': ending_balance,
        'total_debit': total_debit, 'total_credit': total_credit,
    }


# =========================================================================
#  3D. ACB PDF (DA NANG CAP: Dung ty le % chieu rong trang & cat dong chuan xac)
# =========================================================================
_ACB_MONEY_RE = re.compile(r'^\d{1,3}(\.\d{3})+$|^\d{1,3}$')
_ACB_DATE_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')

def extract_acb(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    p0_text = pdf.pages[0].extract_text() or ""
    for line in p0_text.split("\n"):
        line = line.strip()
        if line.startswith("Số tài khoản"):
            m = re.search(r'(\d{5,})', line)
            if m: header_info['account_no'] = m.group(1)
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

    # Tim dong giao dich dau tien tren Trang 1 (co 2 ngay o ben trai: Ngay hieu luc & Ngay giao dich)
    p0 = pdf.pages[0]
    W0, H0 = p0.width, p0.height
    p0_words = p0.extract_words(use_text_flow=False, keep_blank_chars=False)

    col1_dates_p0 = sorted([w for w in p0_words if w['x0'] < 0.13 * W0 and _ACB_DATE_RE.match(w['text'])], key=lambda w: w['top'])
    col2_dates_p0 = [w for w in p0_words if 0.13 * W0 <= w['x0'] < 0.215 * W0 and _ACB_DATE_RE.match(w['text'])]

    first_tx_y = None
    for d1 in col1_dates_p0:
        if any(abs(d2['top'] - d1['top']) < 0.025 * H0 for d2 in col2_dates_p0):
            first_tx_y = d1['top']
            break

    # Xac dinh table_top cua trang 1 ngay duoi tieu de phia tren first_tx_y
    if first_tx_y:
        hdr_above = [w for w in p0_words if (first_tx_y - 0.08 * H0) <= w['top'] < (first_tx_y - 0.006 * H0)]
        p0_table_top = (max([w['bottom'] for w in hdr_above], default=first_tx_y - 20) + 1.5)
    else:
        p0_table_top = 0.38 * H0

    stt_counter = 0
    for pi, page in enumerate(pdf.pages):
        W, H = page.width, page.height
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        table_top = p0_table_top if pi == 0 else 0.015 * H

        # Chan duoi trang: bo qua chu ky ACB va so trang "1 / 2"
        footer_tops = [
            w['top'] for w in words
            if (w['text'].upper() in ('NGÂN', 'ASIA') and w['x0'] > 0.50 * W)
            or (w['text'] == '/' and w['top'] > 0.90 * H)
        ]
        table_bottom = min(footer_tops) - 4 if footer_tops else H - 0.03 * H

        twords = [w for w in words if table_top <= w['top'] < table_bottom]

        # Neo moi dong giao dich la ngay o cot 1 (x0 < 0.13 * W)
        raw_anchors = sorted(
            [w for w in twords if w['x0'] < 0.13 * W and _ACB_DATE_RE.match(w['text'])],
            key=lambda w: (w['top'], w['x0'])
        )
        anchors = []
        for a in raw_anchors:
            if not anchors or abs(a['top'] - anchors[-1]['top']) >= 0.012 * H:
                anchors.append(a)
        if not anchors:
            continue

        # Tim toa do mep phai chuan cua 3 cot tien tu du lieu thuc te tren trang
        # Cot No: x1 trong (0.58W .. 0.73W], Cot Co: (0.73W .. 0.86W], Cot SoDu: > 0.86W
        raw_shapes = [
            s for s in (list(page.rects) + list(page.curves))
            if s['width'] > 0.65 * W and 0.015 * H < s['height'] < 0.25 * H
            and s['top'] >= table_top - 4 and s['bottom'] <= table_bottom + 4
        ]
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
            # Cat dong theo khe ho lon nhat giua cac dong Noi dung (cot 4: 0.265W .. 0.58W)
            cuts = [table_top]
            for i in range(len(anchors) - 1):
                y_a, y_b = anchors[i]['top'], anchors[i + 1]['top']
                desc_w = [w for w in twords if (y_a - 2) <= w['top'] <= (y_b + 2) and 0.265 * W <= w['x0'] < 0.58 * W]
                y_lines = []
                for y in sorted(set(round(w['top'], 1) for w in desc_w)):
                    if not y_lines or y - y_lines[-1] > 3.5:
                        y_lines.append(y)
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
            hl_parts, gd_parts, mota_parts = [], [], []
            no_val, co_val, sodu_val = "0", "0", ""

            for w in sorted(row_words, key=lambda x: (round(x['top']), x['x0'])):
                txt = w['text']
                if w['x0'] < 0.13 * W:
                    hl_parts.append(txt)
                elif w['x0'] < 0.215 * W:
                    gd_parts.append(txt)
                elif w['x0'] < 0.265 * W:
                    continue  # So GD (bo qua cho sach Dien giai)
                elif w['x0'] > 0.57 * W and _ACB_MONEY_RE.match(txt):
                    if w['x1'] <= 0.73 * W:
                        no_val = _vn_money_to_std_str(txt)
                    elif w['x1'] <= 0.86 * W:
                        co_val = _vn_money_to_std_str(txt)
                    else:
                        sodu_val = _vn_money_to_std_str(txt)
                elif w['x0'] < 0.58 * W:
                    mota_parts.append(txt)

            stt_counter += 1
            ngay_gd_str = " ".join(gd_parts)
            ngay_hl_str = " ".join(hl_parts)
            ngay_chuan = normalize_date_ddmmyyyy(ngay_gd_str or ngay_hl_str)

            all_rows.append({
                'STT': str(stt_counter),
                'Ngay': ngay_chuan,
                'MoTa': " ".join(mota_parts).strip(),
                'No': no_val,
                'Co': co_val,
                'SoDu': sodu_val,
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
                'STT': str(stt),
                'Ngay': normalize_date_ddmmyyyy(val_date),
                'MoTa': get(row, "NARR") or pdesc,
                'No': (get(row, "PDAMT") or "0").strip().lstrip("-").strip() or "0",
                'Co': (get(row, "PCAMT") or "0").strip().lstrip("-").strip() or "0",
                'SoDu': get(row, "BOOKED BALANCE"),
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
#  3F. EXCEL / CSV (SACOMBANK + AGRIBANK + CHUNG)
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

            all_rows.append({
                'STT': str(stt), 'Ngay': d_val, 'MoTa': desc,
                'No': f"{abs(amt):,.0f}" if amt < 0 else "0",
                'Co': f"{amt:,.0f}" if amt > 0 else "0",
                'SoDu': f"{to_num(bal):,.0f}" if bal else "",
            })
        if all_rows:
            return {'bank': 'agribank', 'header_info': header_info, 'rows': all_rows,
                    'beginning_balance': f"{to_num(beg_bal):,.0f}" if beg_bal else None,
                    'ending_balance': f"{to_num(end_bal):,.0f}" if end_bal else None,
                    'total_debit': None, 'total_credit': None}

    # 2. SACOMBANK
    vnd_rows = [r for r in rows if "VND" in [str(c).strip() for c in r]]
    if len(vnd_rows) >= 2:
        header_info, all_rows, stt = {}, [], 0
        for r in rows[:25]:
            line = " ".join(str(c) for c in r)
            m = re.search(r'(?:Số|Tài khoản):\s*([0-9A-Za-z]+)', line)
            if m and not header_info.get('account_no'):
                header_info['account_no'] = m.group(1)

        for r in vnd_rows:
            r_str = [str(c).strip() for c in r]
            v_idx = r_str.index("VND")
            desc = r_str[v_idx - 1] if v_idx > 0 else ""

            if "Loại tiền" in desc or "Loại tiền" in " ".join(r_str[:v_idx]):
                continue

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
                all_rows.append({'STT': str(stt), 'Ngay': d_val, 'MoTa': desc, 'No': no_str, 'Co': co_str, 'SoDu': ""})
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
            all_rows.append({'STT': str(stt), 'Ngay': d_val, 'MoTa': texts[0] if texts else "", 'No': nums[0] if nums else "0", 'Co': nums[1] if len(nums) > 1 else "0", 'SoDu': ""})

    return {'bank': 'excel_khac', 'header_info': {}, 'rows': all_rows, 'beginning_balance': None, 'ending_balance': None, 'total_debit': None, 'total_credit': None}


# =========================================================================
#  4. DIEU PHOI, DOI CHIEU & XUAT EXCEL CHUAN 4 COT (NGAY, NO, CO, DIEN GIAI)
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
    while name.lower() in used:
        n += 1
        suffix = f" ({n})"
        name = base[:31 - len(suffix)] + suffix
    used.add(name.lower())
    return name


def build_xlsx(file_paths, out_path, password=None):
    """Xuat file Excel CHUAN 4 COT: Ngay, No, Co, Dien giai"""
    FONT = "Arial"
    header_font = Font(name=FONT, bold=True, color="2C3E50", size=11)
    header_fill = PatternFill("solid", fgColor="DCE6F7")  # Mau xanh pastel nhe
    normal_font = Font(name=FONT, size=10)
    thin = Side(style="thin", color="D0D7E2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_fmt = '#,##0'

    all_results = []
    for p in file_paths:
        res = extract_any(p, password)
        warns = [] if res.get('loi') else validate(res)
        all_results.append((res, warns))

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()

    headers = ["Ngày", "Nợ", "Có", "Diễn giải"]
    widths = [15, 18, 18, 65]

    for res, warns in all_results:
        name = make_sheet_name(res['source_file'], used_names)
        ws = wb.create_sheet(name)

        # Dong 1: Tieu de dung 4 cot
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(1, c, h)
            cell.font, cell.fill, cell.border = header_font, header_fill, border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        if res.get('loi'):
            ws.cell(2, 1, "LỖI")
            ws.cell(2, 4, res['loi'])
        else:
            for r_idx, gd in enumerate(res['rows'], start=2):
                no_num = to_num(gd['No'])
                co_num = to_num(gd['Co'])

                c_ngay = ws.cell(r_idx, 1, gd['Ngay'])
                c_no = ws.cell(r_idx, 2, no_num if no_num != 0 else None)
                c_co = ws.cell(r_idx, 3, co_num if co_num != 0 else None)
                c_dg = ws.cell(r_idx, 4, gd['MoTa'])

                c_ngay.alignment = Alignment(horizontal="center", vertical="top")
                c_no.alignment = Alignment(horizontal="right", vertical="top")
                c_co.alignment = Alignment(horizontal="right", vertical="top")
                c_dg.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

                for cell in (c_ngay, c_no, c_co, c_dg):
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

    wb.save(out_path)
    return all_results


# =========================================================================
#  5. GIAO DIEN WEB PASTEL + NUT NOI 3D + XEM TRUOC 4 COT
# =========================================================================
HTML_UI = """<!DOCTYPE html>
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
            max-width: 1180px;
            margin: 0 auto;
            background: rgba(255, 255, 255, 0.96);
            padding: 32px;
            border-radius: 20px;
            box-shadow: 0 12px 35px rgba(130, 145, 185, 0.16);
            border: 1px solid #eaeef7;
        }
        h1 {
            color: #4b659e;
            margin: 0 0 8px 0;
            font-size: 25px;
            letter-spacing: -0.3px;
        }
        .subtitle {
            color: #6b768c;
            font-size: 14px;
            margin-bottom: 22px;
            line-height: 1.6;
        }
        .upload-zone {
            border: 2.5px dashed #9cb8f3;
            background: #f4f8ff;
            padding: 32px 20px;
            text-align: center;
            border-radius: 16px;
            cursor: pointer;
            margin-bottom: 18px;
            transition: all 0.2s ease;
        }
        .upload-zone:hover {
            background: #e9f1ff;
            border-color: #779be8;
            transform: translateY(-1px);
        }
        .controls {
            display: flex;
            gap: 14px;
            align-items: center;
            flex-wrap: wrap;
            margin-bottom: 16px;
        }
        input[type="password"], input[type="text"] {
            padding: 11px 14px;
            border: 1.5px solid #d9e2f2;
            border-radius: 10px;
            font-size: 14px;
            background: #fbfcff;
            color: #3d485c;
            outline: none;
            transition: border-color 0.2s;
        }
        input:focus { border-color: #8eb0f2; }

        /* HIEU UNG NUT NOI 3D PASTEL */
        .btn {
            padding: 12px 24px;
            border: none;
            border-radius: 12px;
            font-weight: 700;
            font-size: 14px;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            transition: all 0.15s ease;
            transform: translateY(0);
        }
        .btn-primary {
            background: linear-gradient(180deg, #91b3fa 0%, #769ce9 100%);
            color: #ffffff;
            box-shadow: 0 5px 0 #5b7fc9, 0 8px 16px rgba(118, 156, 233, 0.32);
        }
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 7px 0 #5b7fc9, 0 12px 20px rgba(118, 156, 233, 0.38);
        }
        .btn-primary:active {
            transform: translateY(3px);
            box-shadow: 0 2px 0 #5b7fc9, 0 4px 8px rgba(118, 156, 233, 0.25);
        }

        .btn-download {
            background: linear-gradient(180deg, #83ddb8 0%, #63c79e 100%);
            color: #164f38;
            box-shadow: 0 5px 0 #48a67f, 0 8px 16px rgba(99, 199, 158, 0.32);
        }
        .btn-download:hover {
            transform: translateY(-2px);
            box-shadow: 0 7px 0 #48a67f, 0 12px 20px rgba(99, 199, 158, 0.40);
        }
        .btn-download:active {
            transform: translateY(3px);
            box-shadow: 0 2px 0 #48a67f, 0 4px 8px rgba(99, 199, 158, 0.25);
        }

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
        .tab-btn:hover {
            transform: translateY(-1px);
            background: #ebf0fb;
            box-shadow: 0 4px 0 #c9d2e6;
        }
        .tab-btn:active {
            transform: translateY(2px);
            box-shadow: 0 1px 0 #c9d2e6;
        }
        .tab-btn.active {
            background: linear-gradient(180deg, #b8a9f5 0%, #9c8be8 100%);
            color: #ffffff;
            border-color: #8b78db;
            box-shadow: 0 3px 0 #7562c7, 0 5px 10px rgba(156, 139, 232, 0.28);
        }

        #fileList { margin: 10px 0 0 0; font-size: 13px; color: #5b7fc9; font-weight: 600; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; background: #fff; }
        th, td { border: 1px solid #e6ebf5; padding: 10px 12px; text-align: left; }
        th {
            background: #e4ecfc;
            color: #364b75;
            font-weight: 700;
            position: sticky;
            top: 0;
            z-index: 2;
        }
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
        .scroll-box { max-height: 480px; overflow: auto; border: 1px solid #e1e7f2; border-radius: 12px; box-shadow: inset 0 1px 4px rgba(0,0,0,0.03); }
    </style>
</head>
<body>
<div class="container">
    <h1>🌸 Công cụ trích xuất sao kê ngân hàng</h1>
    <div class="subtitle">
        Hỗ trợ: <b>VietinBank, BIDV, LPBank, ACB (PDF)</b> • <b>VPBank (HTML)</b> • <b>Sacombank, Agribank (Excel/CSV)</b>.<br>
        Chuẩn ngày: <b>dd/mm/yyyy</b> • File Excel tải về gồm 4 cột sạch: <b>Ngày, Nợ, Có, Diễn giải</b>.
    </div>

    <div class="upload-zone" onclick="document.getElementById('files').click()">
        <input type="file" id="files" multiple accept=".pdf,.html,.htm,.xlsx,.xls,.csv" style="display:none" onchange="showFiles()">
        <div style="font-size: 16px; font-weight: bold; color: #5b7fc9;">📂 Bấm vào đây để chọn 1 hoặc nhiều file sao kê</div>
        <div id="fileList"></div>
    </div>

    <div class="controls">
        <input type="password" id="pwd" style="width: 320px; max-width: 100%;" placeholder="🔑 Mật khẩu file PDF (nếu có)">
        <button class="btn btn-primary" onclick="uploadAndProcess()">✨ BẮT ĐẦU TRÍCH XUẤT & XEM TRƯỚC</button>
    </div>

    <div id="loader" class="loader">⏳ Đang xử lý và đối chiếu dữ liệu, vui lòng đợi...</div>

    <div id="resultBox" style="display:none;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 14px;">
            <h3 style="margin: 0; color: #4b659e;">📊 Bảng Đối Chiếu Tổng Quan</h3>
            <a id="downloadLink" class="btn btn-download" href="#" download="SaoKe_4Cot_Chuan.xlsx">📥 TẢI FILE EXCEL (4 CỘT)</a>
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
            <h3 style="margin: 0 0 12px 0; color: #4b659e;">🔍 Xem Trước Chi Tiết (4 Cột Chuẩn Excel)</h3>
            <div id="tabs" class="tabs"></div>
            <div class="preview-toolbar">
                <div id="sheetMeta" style="font-size: 13px; color: #5b6982; font-weight: 500;"></div>
                <input type="text" id="searchBox" style="width: 290px;" placeholder="🔎 Tìm theo ngày, số tiền, diễn giải..." oninput="filterPreview()">
            </div>
            <div class="scroll-box">
                <table>
                    <thead>
                        <tr>
                            <th class="text-center" style="width: 115px;">Ngày</th>
                            <th class="text-right" style="width: 145px;">Nợ</th>
                            <th class="text-right" style="width: 145px;">Có</th>
                            <th>Diễn giải</th>
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
                <td class="text-right">${r.count}</td><td class="text-right debit">${r.debit}</td>
                <td class="text-right credit">${r.credit}</td><td><span class="${cls}">${r.status}</span></td>`;
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
        `Ngân hàng: <b>${s.bank}</b> • Số TK: <b>${s.account || '—'}</b> • Tổng số giao dịch: <b>${s.rows.length}</b>`;
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
        (r.Co || '').toLowerCase().includes(kw)
    );
    renderPreviewRows(filtered);
}

function renderPreviewRows(rows) {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';
    if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="4" class="text-center" style="color:#888;padding:24px;">Không có dữ liệu giao dịch phù hợp</td></tr>';
        return;
    }
    rows.forEach(r => {
        let tr = document.createElement('tr');
        let noDisp = (r.No && r.No !== '0') ? r.No : '';
        let coDisp = (r.Co && r.Co !== '0') ? r.Co : '';
        tr.innerHTML = `
            <td class="text-center"><b>${r.Ngay}</b></td>
            <td class="text-right debit">${noDisp}</td>
            <td class="text-right credit">${coDisp}</td>
            <td>${r.MoTa}</td>
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

        out_xlsx = os.path.join(tmp_dir, "SaoKe_4Cot_Chuan.xlsx")
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
