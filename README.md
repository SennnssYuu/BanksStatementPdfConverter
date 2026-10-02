# BanksStatementPdfConverter

Convert locked bank statement PDFs to Excel, offline. A Thai/English desktop app for Windows.

แปลงไฟล์ PDF ใบแจ้งยอดบัญชี (รวมถึงไฟล์ที่มีรหัสผ่าน) เป็นไฟล์ Excel ทำงานในเครื่องเท่านั้น ไม่ส่งข้อมูลออกออนไลน์

## Features

- **Opens password-protected PDFs.** The password is typed into the app, used once, and never saved.
- **Built for KBank (Kasikornbank) statements.** Reads the main table (วันที่ · เวลา · รายการ · ถอนเงิน / ฝากเงิน · ยอดคงเหลือ · ช่องทาง · รายละเอียด), including:
  - Withdrawal and deposit split into separate columns, decided by whether the balance went up or down
  - Descriptions that wrap onto two lines (e.g. "รับเงินจากการขายด้วย Thai / QR Payment")
  - Statements with many pages
  - The account details box at the top; the account number is added to every row
- **Checks its own work.** Compares the extracted transactions with the statement's totals (รวมถอนเงิน, รวมฝากเงิน, ยอดยกไป) and checks the running balance on every row. Problem rows are highlighted red.
- **Never drops a line silently.** Anything in the table it can't read is listed on a separate sheet.
- **Thai / English interface**, switchable from the top-right corner.
- **Fully offline.** No network code. Files never leave the computer.
- Other banks' statements go through a general-purpose reader (see [Limitations](#limitations)).

## Download

Get **`Statement to Excel.exe`** from [Releases](../../releases). It's a single file that runs on Windows 10/11 without Python installed. Copy it to any PC (e.g. by USB) and double-click it.

Because the `.exe` isn't digitally signed, Windows may show "Windows protected your PC" the first time. Click **More info → Run anyway**.

## How to use

1. **Choose your statement PDF**: click **เลือกไฟล์… / Browse…**
2. **Password**: type the PDF password (only shown if the file is locked)
3. **Save Excel file as**: defaults to the same folder as the PDF
4. Click **แปลงเป็น Excel / Convert to Excel**

When it finishes, a green ✓ means every row was read and the totals match the statement. A yellow message tells you what to check.

## The Excel file

| Sheet | Contents |
|---|---|
| **Transactions** | Account No. · Date · Time · Description · Withdrawal · Deposit · Balance · Channel · Details · Page · Check. Dates are real Excel dates, the header has filters, and there's a total row. |
| **Account Info** | Account name, account number, reference, period, branch, opening/closing balance, plus a check table: statement totals vs extracted totals. |
| **Skipped lines** | Only appears if some lines in the table couldn't be read. |
| **Raw text** | All text from the PDF, for checking. |

## Privacy

The app has no internet code and opens no network connections while converting. Things to watch outside the app:

- Keep statements in a folder that **doesn't sync** to OneDrive or Google Drive (on many PCs, Desktop and Documents sync to OneDrive).
- This repository's `.gitignore` excludes `*.pdf` and `*.xlsx`, so statements can't be committed by accident.

## Run from source

Requires Python 3.10+.

```bash
pip install -r requirements.txt
python app.py
```

Command-line version (asks for the password if needed):

```bash
python statement_to_excel.py statement.pdf -o output.xlsx
```

### Build the .exe

Double-click `build_exe.bat`, or run it from a terminal. The result is `dist\Statement to Excel.exe`.

### Troubleshooting a statement

If a row comes out wrong, write a layout report for the lines around it:

```bash
python statement_to_excel.py statement.pdf --debug "04-09-26"
```

This creates `debug.txt` next to the PDF, showing where each word sits on the page and which column it was assigned to. **Blank out names and account numbers before sharing it.**

`unlock_copy.py` saves a password-free copy of a PDF for testing. Delete the copy afterwards.

## Project structure

| File | Purpose |
|---|---|
| `app.py` | Desktop window (Thai/English) |
| `statement_to_excel.py` | Conversion entry point, command line, general-purpose reader |
| `kbank.py` | KBank statement reader, checks and Excel layout |
| `unlock_copy.py` | Save an unlocked copy of a PDF (for debugging) |
| `build_exe.bat` | Build the single-file `.exe` with PyInstaller |
| `Statement to Excel.bat` | Start the app from source with a double-click |
| `icon.ico` | App icon |

## Limitations

- Only KBank statements have a dedicated reader. Other banks get a general reader that finds dated lines but doesn't separate withdrawals and deposits; check its output carefully.
- Scanned statements (images of pages, with no selectable text) can't be read.
- Tested on Windows 10.
