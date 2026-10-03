"""
Build a SAMPLE statement PDF for testing the converter.

The data is transcribed from a bank's publicly published *sample* statement
(fictional account "นายกสิกร หมั่นเพียร"): 22 rows, 7 withdrawals, 14 deposits.
The PDF has real, selectable text in the same 7-column layout as a KBank
savings statement, no bank logo or branding, and a clear "SAMPLE" label.

Usage (Windows, needs the Tahoma font that ships with Windows):
    pip install reportlab
    python samples/make_sample_pdf.py
    python statement_to_excel.py samples/sample-statement.pdf -o samples/sample-statement.xlsx

Output goes next to this script. *.pdf / *.xlsx are git-ignored, so only this
script is committed. Expected result: KBank format, 22 transactions, all checks OK.
"""
import re
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

OUT = Path(__file__).with_name("sample-statement.pdf")

# --------------------------------------------------------------- sample data
# Transcribed from KBank's public SAMPLE statement image (fictional account "นายกสิกร หมั่นเพียร").
# (date, time, description, withdrawal, deposit, balance, channel, details)
rows = [
    ("05-01-20", "", "ยอดยกมา", None, None, "1,520,000.00", "", ""),
    ("05-01-20", "10:25", "รับโอนเงิน", None, "28,000.00", "1,548,000.00", "Internet/Mobile TMB", "จาก TMB X0375 นายอนุรักษ์ ก++"),
    ("05-01-20", "12:28", "รับโอนเงิน", None, "15,000.00", "1,563,000.00", "K PLUS SME", "จาก X3233 นาย พงศ์สวัสด++"),
    ("05-01-20", "14:45", "รับโอนเงินอัตโนมัติ", None, "10,200.00", "1,573,200.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "จาก SMART SCB X2943 จุนาพงศ์ ธนาป++"),
    ("05-01-20", "16:50", "รับเงินจากการขายด้วย Thai QR Payment", None, "70,000.00", "1,643,200.00", "EDC/K PLUS SHOP", "จาก 401010390952001 อเมซอน-ทรรศพร"),
    ("06-01-20", "09:25", "โอนเงิน", "80,000.00", None, "1,563,200.00", "K-Cash Connect Plus", "โอนไป AC Link X8899 ดำรง บัวผันกา++"),
    ("06-01-20", "10:30", "โอนเงิน LMS", "50,000.00", None, "1,513,200.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "ไป X6687 ฝันวงศ์ สุพัน++"),
    ("06-01-20", "11:45", "ชำระค่าไฟฟ้า", "36,500.00", None, "1,476,700.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "เพื่อชำระ Ref x3510 การไฟฟ้านครหลวง"),
    ("06-01-20", "12:50", "รับโอนเงิน", None, "48,000.00", "1,524,700.00", "K PLUS SME", "จาก x2233 นาย นิยม หรรษ++"),
    ("06-01-20", "14:55", "ชำระค่าโทรศัพท์", "80,000.00", None, "1,444,700.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "เพื่อชำระ Ref x8899 บริษัท ทีโอที จำกัด(มหาชน)"),
    ("07-01-20", "10:58", "รับชำระเงิน", None, "68,000.00", "1,512,700.00", "ATM", "Bill Payment"),
    ("07-01-20", "12:15", "รับโอนเงิน", None, "33,800.00", "1,546,500.00", "K PLUS SME", "จาก x8115 นาย อภินันท์ ห++"),
    ("07-01-20", "16:12", "โอนเงิน", "15,036.00", None, "1,531,464.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "โอนไป SCB X0206 นางสาว อัมริน++"),
    ("07-01-20", "23:30", "รับเงินจากการขาย เติมจำนวน/ผ่อนชำระ/คะแนนสะสม", None, "98,700.00", "1,630,164.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "จาก 123456789 บริษัท กขค จำกัด"),
    ("08-01-20", "14:00", "โอนเงิน LMS", "120,000.00", None, "1,510,164.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "ไป X1234 บจก. กิจเจริญ++"),
    ("08-01-20", "15:38", "รับโอนเงินอัตโนมัติ", None, "278,000.00", "1,788,164.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "จาก SMART BBL X1234 นางสมใจ แก้วต++"),
    ("08-01-20", "16:20", "ฝากด้วยเช็ค", None, "40,000.00", "1,828,164.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "KTB 123 เช็คเลขที่ 04034055"),
    ("08-01-20", "17:45", "ชำระค่าน้ำประปา", "12,305.50", None, "1,815,858.50", "โอนเข้า/หักบัญชีอัตโนมัติ", "เพื่อชำระ Ref x3511 การประปานครหลวง"),
    ("09-01-20", "10:50", "รับโอนเงิน", None, "18,500.50", "1,834,359.00", "Internet/Mobile SCB", "จาก SCB X0375 MISS JIRAPORN++"),
    ("09-01-20", "13:40", "รับเงินจากการขายด้วย Thai QR Payment", None, "32,500.00", "1,866,859.00", "EDC/K PLUS SHOP", "จาก 401010390952022 บริษัท เจริญผล"),
    ("10-01-20", "11:20", "รับโอนเงิน", None, "30,400.00", "1,897,259.00", "K-Cash Connect Plus", "จาก X1234 นายสมบูรณ์ บุ++"),
    ("10-01-20", "15:15", "รับโอนเงินบาทเนต", None, "124,000.00", "2,021,259.00", "โอนเข้า/หักบัญชีอัตโนมัติ", "จาก KRUNG THAI BANK PCL. VICHAYA TANAS++"),
]


def num(s):
    return float(s.replace(",", "")) if s else 0.0


# Self-check the transcription before building anything.
bal = num(rows[0][5])
wd_total = dp_total = 0.0
n_wd = n_dp = 0
for r in rows[1:]:
    bal += num(r[4]) - num(r[3])
    wd_total += num(r[3])
    dp_total += num(r[4])
    n_wd += bool(r[3])
    n_dp += bool(r[4])
    assert abs(bal - num(r[5])) < 0.005, r
assert (n_wd, round(wd_total, 2), n_dp, round(dp_total, 2)) == (7, 393841.5, 14, 895100.5)
print("sample data checks out:", n_wd, "withdrawals", wd_total, "|", n_dp, "deposits", dp_total, "| closing", bal)


# --------------------------------------------------------------- PDF
pdfmetrics.registerFont(TTFont("Tahoma", "C:/Windows/Fonts/tahoma.ttf"))
pdfmetrics.registerFont(TTFont("Tahoma-Bold", "C:/Windows/Fonts/tahomabd.ttf"))
W, H = A4
FS = 7.2  # table font size
LINE = 9.2  # wrapped line spacing
ROW_GAP = 4.6

# column edges (x) for: date, time, desc, amount, balance, channel, details
X = [28, 76, 112, 232, 312, 378, 462, 567]


def tw(text, font="Tahoma", size=FS):
    return pdfmetrics.stringWidth(text, font, size)


def wrap(text, width, font="Tahoma", size=FS):
    """Wrap at spaces, or right after '/' (the statement wraps there too)."""
    if not text:
        return [""]
    tokens = re.findall(r"[^ /]+/?|/| ", text)
    lines, cur = [], ""
    for tok in tokens:
        cand = cur + tok
        if cur and tw(cand.rstrip(), font, size) > width:
            lines.append(cur.rstrip())
            cur = tok.lstrip()
        else:
            cur = cand
    if cur.strip():
        lines.append(cur.rstrip())
    return lines


c = canvas.Canvas(str(OUT), pagesize=A4)
c.setTitle("Sample statement (test document)")
c.setAuthor("Statement to Excel test data")

# --- SAMPLE label
c.setFillColorRGB(0.93, 0.93, 0.93)
c.rect(28, H - 52, 74, 24, stroke=0, fill=1)
c.setFillColorRGB(0.8, 0, 0)
c.setFont("Tahoma-Bold", 15)
c.drawString(36, H - 45, "ตัวอย่าง")
c.setFillColorRGB(0.4, 0.4, 0.4)
c.setFont("Tahoma", 6.5)
c.drawString(110, H - 43, "SAMPLE ONLY · recreated from a public example statement for testing · not a real account")

# --- title
c.setFillColorRGB(0, 0, 0)
c.setFont("Tahoma-Bold", 12)
c.drawString(28, H - 76, "รายการเดินบัญชีเงินฝากออมทรัพย์ (มีรายละเอียด)")
c.setFont("Tahoma", 7)
c.drawString(28, H - 87, "DEPOSIT STATEMENT OF SAVING ACCOUNT")

# --- left block
c.setFont("Tahoma", 8)
c.drawString(28, H - 108, "หน้าที่ 1/1 (0001)")
c.drawString(28, H - 124, "ชื่อบัญชี นายกสิกร หมั่นเพียร")
c.drawString(40, H - 140, "888 ถนนราษฎร์บูรณะ แขวงราษฎร์บูรณะ เขตราษฎร์บูรณะ กทม. 10140")

# --- account box (right)
bx, bw, lw = 352, 215, 92
rh = 16


def box(rows_, top):
    y = top
    for label, value, right in rows_:
        c.rect(bx, y - rh, bw, rh, stroke=1, fill=0)
        c.line(bx + lw, y, bx + lw, y - rh)
        c.setFont("Tahoma-Bold", 7.5)
        c.drawString(bx + 4, y - rh + 5, label)
        c.setFont("Tahoma", 7.5)
        if right:
            c.drawRightString(bx + bw - 5, y - rh + 5, value)
        else:
            c.drawString(bx + lw + 5, y - rh + 5, value)
        y -= rh
    return y


c.setLineWidth(0.6)
y = box([
    ("เลขที่อ้างอิง", "20013116520037461269", False),
    ("เลขที่บัญชีเงินฝาก", "001-2-00003-0", False),
    ("รอบระหว่างวันที่", "01/01/2020 - 31/03/2020", False),
    ("สาขาเจ้าของบัญชี", "สาขาสำนักสีลม", False),
], H - 100)
box([
    ("ยอดยกไป", "2,021,259.00", True),
    ("รวมถอนเงิน 7 รายการ", "393,841.50", True),
    ("รวมฝากเงิน 14 รายการ", "895,100.50", True),
], y - 6)

# --- transactions table
top = H - 236
head_h = 26
c.setLineWidth(0.7)
c.line(X[0], top, X[-1], top)
c.line(X[0], top - head_h, X[-1], top - head_h)
headers = ["วันที่", ("เวลา/", "วันที่มีผล"), "รายการ", "ถอนเงิน / ฝากเงิน", "ยอดคงเหลือ", "ช่องทาง", "รายละเอียด"]
c.setFont("Tahoma-Bold", 7.5)
for i, h in enumerate(headers):
    cx = (X[i] + X[i + 1]) / 2
    if isinstance(h, tuple):
        c.drawCentredString(cx, top - 11, h[0])
        c.drawCentredString(cx, top - 21, h[1])
    else:
        c.drawCentredString(cx, top - 16, h)

y = top - head_h - 12
c.setFont("Tahoma", FS)
pad = 3
for d, t, desc, wd, dp, bl, ch, det in rows:
    desc_l = wrap(desc, X[3] - X[2] - 2 * pad)
    ch_l = wrap(ch, X[6] - X[5] - 2 * pad)
    det_l = wrap(det, X[7] - X[6] - 2 * pad)
    n = max(len(desc_l), len(ch_l), len(det_l))
    c.drawString(X[0] + pad, y, d)
    if t:
        c.drawCentredString((X[1] + X[2]) / 2, y, t)
    for k, s in enumerate(desc_l):
        c.drawString(X[2] + pad, y - k * LINE, s)
    if wd:
        c.drawString(X[3] + pad, y, wd)  # withdrawals sit on the left of the amount column
    if dp:
        c.drawRightString(X[4] - pad, y, dp)  # deposits on the right
    c.drawRightString(X[5] - pad, y, bl)
    for k, s in enumerate(ch_l):
        c.drawString(X[5] + pad, y - k * LINE, s)
    for k, s in enumerate(det_l):
        c.drawString(X[6] + pad, y - k * LINE, s)
    y -= (n - 1) * LINE + LINE + ROW_GAP

bottom = y + 4
c.line(X[0], bottom, X[-1], bottom)
for x in X:
    c.line(x, top, x, bottom)

c.setFont("Tahoma", 6.5)
c.setFillColorRGB(0.35, 0.35, 0.35)
c.drawString(28, bottom - 18, "Test document for the Statement to Excel converter. All names and numbers come from a published sample statement.")
c.showPage()
c.save()
print(f"Wrote {OUT}")
print(f"Try it:  python statement_to_excel.py \"{OUT}\" -o \"{OUT.with_suffix('.xlsx')}\"")
