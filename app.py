# -*- coding: utf-8 -*-
"""
CONG CU TRICH XUAT SAO KE NGAN HANG
- PDF: VietinBank, BIDV, LPBank (7 cot & 10 cot), ACB (fix nhan dien cot tien theo mep phai x1)
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
import unicodedata
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
    t = unicodedata.normalize("NFC", text or "")
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
#  3D. ACB PDF (Phan loai 3 cot tien bang mep phai x1 tu tieu de)
# =========================================================================
_ACB_MONEY_RE = re.compile(r'^\d{1,3}([.,]\d{3})+$|^\d{1,3}$')
_ACB_DATE_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')

def extract_acb(pdf):
    header_info, all_rows = {}, []
    beginning_balance = ending_balance = total_debit = total_credit = None

    p0 = pdf.pages[0]
    W0, H0 = p0.width, p0.height
    p0_text = unicodedata.normalize("NFC", p0.extract_text() or "")

    for line in p0_text.split("\n"):
        line = line.strip()
        if line.startswith("Số tài khoản"):
            m = re.search(r'(\d{5,})', line)
            if m: header_info['account_no'] = m.group(1)
        elif line.startswith("Số dư đầu"):
            m = re.search(r'([\d.,]+)$', line)
            if m: beginning_balance = _vn_money_to_std_str(m.group(1))
        elif line.startswith("Tổng tiền rút ra"):
            m = re.search(r'([\d.,]+)$', line)
            if m: total_debit = _vn_money_to_std_str(m.group(1))
        elif line.startswith("Tổng tiền gửi vào"):
            m = re.search(r'([\d.,]+)$', line)
            if m: total_credit = _vn_money_to_std_str(m.group(1))
        elif line.startswith("Số dư cuối"):
            m = re.search(r'([\d.,]+)$', line)
            if m: ending_balance = _vn_money_to_std_str(m.group(1))

    p0_words = p0.extract_words(use_text_flow=False, keep_blank_chars=False)
    for w in p0_words:
        w['text'] = unicodedata.normalize("NFC", w['text'])

    col1_dates_p0 = sorted([w for w in p0_words if w['x0'] < 0.13 * W0 and _ACB_DATE_RE.match(w['text'])], key=lambda w: w['top'])
    col2_dates_p0 = [w for w in p0_words if 0.13 * W0 <= w['x0'] < 0.22 * W0 and _ACB_DATE_RE.match(w['text'])]

    first_tx_y = None
    for d1 in col1_dates_p0:
        if any(abs(d2['top'] - d1['top']) < 0.025 * H0 for d2 in col2_dates_p0):
            first_tx_y = d1['top']
            break

    if first_tx_y:
        hdr_above = [w for w in p0_words if (first_tx_y - 0.08 * H0) <= w['top'] < (first_tx_y - 0.005 * H0)]
        p0_table_top = (max([w['bottom'] for w in hdr_above], default=first_tx_y - 20) + 1.5)
    else:
        hdr_above = []
        p0_table_top = 0.38 * H0

    x1_rut = max([w['x1'] for w in hdr_above if w['text'].lower() in ('ra', 'rút', 'debit', 'withdrawal') and w['x0'] > 0.45 * W0], default=0.674 * W0)
    x1_gui = max([w['x1'] for w in hdr_above if w['text'].lower() in ('vào', 'gửi', 'credit', 'deposit') and w['x0'] > x1_rut], default=0.805 * W0)
    x1_du = max([w['x1'] for w in hdr_above if w['text'].lower() in ('dư', 'balance') and w['x0'] > x1_gui], default=0.935 * W0)

    money_min_x1 = x1_rut - 0.05 * W0
    mid_rut_gui = (x1_rut + x1_gui) / 2
    mid_gui_du = (x1_gui + x1_du) / 2

    stt_counter = 0
    for pi, page in enumerate(pdf.pages):
        W, H = page.width, page.height
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        for w in words:
            w['text'] = unicodedata.normalize("NFC", w['text'])

        table_top = p0_table_top if pi == 0 else 0.015 * H

        footer_tops = [
            w['top'] for w in words
            if (w['text'].upper() in ('NGÂN', 'ASIA') and w['x0'] > 0.50 * W)
            or (w['text'] == '/' and w['top'] > 0.90 * H)
        ]
        table_bottom = min(footer_tops) - 4 if footer_tops else H - 0.03 * H
        twords = [w for w in words if table_top <= w['top'] < table_bottom]

        scale = W / W0 if W0 else 1.0
        cur_money_min_x1 = money_min_x1 * scale
        cur_mid_rg = mid_rut_gui * scale
        cur_mid_gd = mid_gui_du * scale

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
            cuts = [table_top]
            for i in range(len(anchors) - 1):
                y_a, y_b = anchors[i]['top'], anchors[i + 1]['top']
                desc_w = [w for w in twords if (y_a - 2) <= w['top'] <= (y_b + 2) and 0.26 * W <= w['x0'] < (x1_rut * scale - 0.09 * W)]
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
                elif w['x0'] < 0.265 * W and re.match(r'^\d{1,6}$', txt):
                    continue
                elif w['x1'] >= cur_money_min_x1 and _ACB_MONEY_RE.match(txt):
                    if w['x1'] <= cur_mid_rg:
                        no_val = _vn_money_to_std_str(txt)
                    elif w['x1'] <= cur_mid_gd:
                        co_val = _vn_money_to_std_str(txt)
                    else:
                        sodu_val = _vn_money_to_std_str(txt)
                else:
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
#  4. DIEU PHOI, DOI CHIEU & XUAT EXCEL CHUAN 4 COT: NGAY, NO, CO, DIEN GIAI
# ============
