"""
Statement to Excel - desktop window (Thai / English).

Pick a statement PDF with File Explorer, type its password (if it has one),
and get an Excel file. Everything runs locally; nothing is uploaded.

Run:  python app.py   (or double-click "Statement to Excel.bat")
"""

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

from statement_to_excel import PasswordError, convert, is_locked

# Colors
BG = "#F3F5F8"
CARD = "#FFFFFF"
BORDER = "#DDE2EA"
TEXT = "#1B2430"
MUTED = "#5E6B7D"
PRIMARY = "#1F4E78"
PRIMARY_HOVER = "#173B5C"
SUCCESS = "#1E7B45"
ERROR = "#B42318"
WARN = "#9A6700"

SETTINGS = Path(os.environ.get("APPDATA", Path.home())) / "StatementToExcel" / "settings.json"


def _plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


STRINGS = {
    "th": {
        "window": "แปลงใบแจ้งยอดเป็น Excel",
        "title": "แปลงใบแจ้งยอดเป็น Excel",
        # Thai has no spaces between words, so Tk can't wrap it: break lines by hand
        "subtitle": "แปลงไฟล์ PDF ใบแจ้งยอดบัญชีเป็นไฟล์ Excel\n"
                    "ทำงานในเครื่องนี้เท่านั้น ไฟล์ของคุณจะไม่ถูกส่งออกไปที่ใด",
        "step1": "เลือกไฟล์ PDF ใบแจ้งยอด",
        "browse": "เลือกไฟล์…",
        "no_file": "ยังไม่ได้เลือกไฟล์  ·  Ctrl+O",
        "bad_pdf": "เปิดไฟล์นี้เป็น PDF ไม่ได้ กรุณาเลือกไฟล์อื่น",
        "locked": "มีรหัสผ่าน",
        "not_locked": "ไม่มีรหัสผ่าน",
        "step2": "รหัสผ่าน",
        "show": "แสดง",
        "pw_choose_first": "กรุณาเลือกไฟล์ PDF ก่อน",
        "pw_not_locked": "ไฟล์นี้ไม่มีรหัสผ่าน ไม่ต้องกรอก",
        "pw_enter": "กรอกรหัสผ่านของไฟล์ PDF นี้",
        "pw_ready": "กด Enter หรือคลิก “แปลงเป็น Excel”",
        "pw_wrong": "รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่",
        "step3": "บันทึกไฟล์ Excel เป็น",
        "change": "เปลี่ยน…",
        "out_hint": "ค่าเริ่มต้นคือโฟลเดอร์เดียวกับไฟล์ PDF",
        "convert": "แปลงเป็น Excel",
        "converting": "กำลังแปลง…",
        "opening": "กำลังเปิดไฟล์ PDF…",
        "reading": "กำลังอ่านหน้า {done} จาก {total}…",
        "pick_title": "เลือกไฟล์ PDF ใบแจ้งยอด",
        "pdf_files": "ไฟล์ PDF",
        "all_files": "ทุกไฟล์",
        "save_title": "บันทึกไฟล์ Excel เป็น",
        "excel_files": "ไฟล์ Excel",
        "replace_title": "แทนที่ไฟล์เดิม?",
        "replace_msg": "มีไฟล์ {name} อยู่แล้ว\n\nต้องการบันทึกทับหรือไม่?",
        "err_save": "บันทึกไฟล์ Excel ไม่ได้\nถ้าไฟล์นี้เปิดอยู่ใน Excel ให้ปิดก่อนแล้วลองใหม่",
        "err_other": "เกิดข้อผิดพลาด: {exc}",
        "done": "เสร็จแล้ว  {found}",
        "done_check": "เสร็จแล้ว กรุณาตรวจสอบ  {found}",
        "found": lambda n, pages: f"พบ {n} รายการ จาก {pages} หน้า",
        "account": "บัญชี KBank {acct}",
        "acct_missing": "(ไม่พบเลขที่บัญชี)",
        "skipped": "มี {n} บรรทัดในตารางที่อ่านไม่ได้ ดูได้ในชีต “Skipped lines”",
        "mismatch": "ไม่ตรงกับใบแจ้งยอด: {items}\nดูรายละเอียดในชีต “Account Info” "
                    "(บรรทัดที่มีปัญหาจะเป็นสีแดง)",
        "unknown": "ตรวจยอดคงเหลือทุกบรรทัดแล้ว\nแต่ไม่พบยอดรวมในใบแจ้งยอดสำหรับเปรียบเทียบ",
        "all_ok": "✓ ยอดรวมและยอดยกไปตรงกับใบแจ้งยอด",
        "none_found": "ไม่พบรายการที่มีวันที่ ลองดูชีต “Tables” และ “Raw text”\n"
                      "ถ้าว่างเปล่า ไฟล์ PDF อาจเป็นภาพสแกน",
        "saved": "บันทึกไว้ที่ {path}",
        "open_xlsx": "เปิดไฟล์ Excel",
        "show_folder": "เปิดโฟลเดอร์",
        "another": "แปลงไฟล์อื่น",
    },
    "en": {
        "window": "Statement to Excel",
        "title": "Statement to Excel",
        "subtitle": "Turn a statement PDF into an Excel file. "
                    "Runs offline, your file never leaves this computer.",
        "step1": "Choose your statement PDF",
        "browse": "Browse…",
        "no_file": "No file selected  ·  Ctrl+O",
        "bad_pdf": "This file couldn't be opened as a PDF. Choose another file.",
        "locked": "Password protected",
        "not_locked": "Not locked",
        "step2": "Password",
        "show": "Show",
        "pw_choose_first": "Choose a PDF first.",
        "pw_not_locked": "This PDF isn't locked, no password needed.",
        "pw_enter": "Enter the password for this PDF.",
        "pw_ready": "Press Enter or click Convert.",
        "pw_wrong": "Wrong password. Try again.",
        "step3": "Save Excel file as",
        "change": "Change…",
        "out_hint": "Defaults to the same folder as the PDF.",
        "convert": "Convert to Excel",
        "converting": "Converting…",
        "opening": "Opening PDF…",
        "reading": "Reading page {done} of {total}…",
        "pick_title": "Choose a statement PDF",
        "pdf_files": "PDF files",
        "all_files": "All files",
        "save_title": "Save Excel file as",
        "excel_files": "Excel workbook",
        "replace_title": "Replace file?",
        "replace_msg": "{name} already exists.\n\nReplace it?",
        "err_save": "Couldn't save the Excel file. If it's open in Excel, close it and try again.",
        "err_other": "Something went wrong: {exc}",
        "done": "Done. {found}",
        "done_check": "Done, please check. {found}",
        "found": lambda n, pages: f"{_plural(n, 'transaction')} found on {_plural(pages, 'page')}",
        "account": "KBank account {acct}",
        "acct_missing": "(account number not found)",
        "skipped": "{n} line(s) in the table couldn't be read. "
                   "They're listed on the 'Skipped lines' sheet.",
        "mismatch": "Doesn't match the statement: {items}. "
                    "See the 'Account Info' sheet; rows with problems are red.",
        "unknown": "Running balance checked on every row. "
                   "Statement totals weren't found to compare.",
        "all_ok": "✓ Totals and closing balance match the statement.",
        "none_found": "No dated lines were found. Check the 'Tables' and 'Raw text' sheets. "
                      "If they're empty, the PDF may be a scanned image.",
        "saved": "Saved to {path}",
        "open_xlsx": "Open Excel file",
        "show_folder": "Show in folder",
        "another": "Convert another",
    },
}


def human_size(n):
    for unit in ("bytes", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "bytes" else f"{n:.1f} {unit}"
        n /= 1024


def load_lang():
    try:
        lang = json.loads(SETTINGS.read_text(encoding="utf-8")).get("lang")
        return lang if lang in STRINGS else "th"
    except Exception:
        return "th"


def save_lang(lang):
    try:
        SETTINGS.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS.write_text(json.dumps({"lang": lang}), encoding="utf-8")
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.configure(bg=BG)
        self.minsize(580, 0)
        self.resizable(True, False)

        families = set(tkfont.families(self))
        self.font = next((f for f in ("Leelawadee UI", "Tahoma", "Segoe UI") if f in families),
                         "TkDefaultFont")
        self.lang = load_lang()

        self.pdf_path = None
        self.out_path = None
        self.locked = False
        self.pdf_error = False
        self.busy = False
        self.pw_wrong = False
        self.status = ("", None)    # (string key or text, kwargs)
        self.status_color = MUTED
        self.last_result = None     # stats of the last finished conversion
        self.events = queue.Queue()

        self.pdf_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.pw_var = tk.StringVar()
        self.show_pw = tk.BooleanVar(value=False)
        self.pw_var.trace_add("write", lambda *_: self._on_password_typed())

        self._init_styles()
        self.page = None
        self._build()

        self.bind("<Return>", lambda e: self.start_convert())
        self.bind("<Control-o>", lambda e: self.choose_pdf())
        self.after(50, self._poll_events)

    def t(self, key, **kw):
        s = STRINGS[self.lang][key]
        return s(**kw) if callable(s) else s.format(**kw)

    # ---------- look ----------
    def _init_styles(self):
        F = self.font
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", font=(F, 10), background=BG, foreground=TEXT)
        s.configure("Card.TFrame", background=CARD)
        s.configure("Card.TLabel", background=CARD)
        s.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=(F, 9))
        s.configure("Step.TLabel", background=CARD, font=(F, 11, "bold"))
        s.configure("Title.TLabel", background=BG, font=(F, 18, "bold"))
        s.configure("Sub.TLabel", background=BG, foreground=MUTED)
        s.configure("Status.TLabel", background=BG, foreground=MUTED)
        s.configure("Card.TCheckbutton", background=CARD)
        s.map("Card.TCheckbutton", background=[("active", CARD)])
        s.configure("TEntry", padding=6, fieldbackground=CARD, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER)
        s.map("TEntry", bordercolor=[("focus", PRIMARY)], lightcolor=[("focus", PRIMARY)],
              fieldbackground=[("readonly", "#F7F8FA"), ("disabled", "#EEF0F3")])
        s.configure("TButton", padding=(14, 6), background="#E9EDF2", bordercolor=BORDER,
                    lightcolor="#E9EDF2", darkcolor="#E9EDF2")
        s.map("TButton", background=[("active", "#DCE2EA"), ("disabled", "#F0F2F5")])
        s.configure("Primary.TButton", font=(F, 11, "bold"), padding=(20, 10),
                    background=PRIMARY, foreground="white", bordercolor=PRIMARY,
                    lightcolor=PRIMARY, darkcolor=PRIMARY)
        s.map("Primary.TButton",
              background=[("disabled", "#A9B7C6"), ("active", PRIMARY_HOVER)],
              foreground=[("disabled", "#EEF2F6")])
        s.configure("Horizontal.TProgressbar", troughcolor="#E3E8EF", background=PRIMARY,
                    bordercolor=BG, lightcolor=PRIMARY, darkcolor=PRIMARY, thickness=8)
        # Language switch
        s.configure("Lang.TButton", padding=(10, 3), font=(F, 9), background=BG,
                    foreground=MUTED, bordercolor=BORDER, lightcolor=BG, darkcolor=BG)
        s.map("Lang.TButton", background=[("active", "#E3E8EF")])
        s.configure("LangOn.TButton", padding=(10, 3), font=(F, 9, "bold"), background=PRIMARY,
                    foreground="white", bordercolor=PRIMARY, lightcolor=PRIMARY, darkcolor=PRIMARY)
        s.map("LangOn.TButton", background=[("active", PRIMARY_HOVER)])

    def _card(self, parent, step, title):
        outer = tk.Frame(parent, bg=BORDER)  # 1px border
        outer.pack(fill="x", pady=(0, 12))
        card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        card.pack(fill="both", padx=1, pady=1)
        ttk.Label(card, text=f"{step}   {title}", style="Step.TLabel").pack(anchor="w")
        return outer, card

    # ---------- layout ----------
    def _build(self):
        """(Re)build the whole window in the current language, keeping the current state."""
        if self.page is not None:
            self.page.destroy()
        self.title(self.t("window"))
        root = self.page = ttk.Frame(self, padding=24)
        root.pack(fill="both", expand=True)

        top = ttk.Frame(root)
        top.pack(fill="x")
        ttk.Label(top, text=self.t("title"), style="Title.TLabel").pack(side="left", anchor="w")
        switch = ttk.Frame(top)
        switch.pack(side="right", anchor="n")
        for code, label in (("th", "ไทย"), ("en", "EN")):
            ttk.Button(switch, text=label, width=4,
                       style="LangOn.TButton" if code == self.lang else "Lang.TButton",
                       command=lambda c=code: self.set_lang(c)).pack(side="left")
        ttk.Label(root, text=self.t("subtitle"), style="Sub.TLabel",
                  wraplength=520, justify="left").pack(anchor="w", pady=(2, 18))

        # Step 1 - PDF
        _, c1 = self._card(root, "1", self.t("step1"))
        row = ttk.Frame(c1, style="Card.TFrame")
        row.pack(fill="x", pady=(10, 0))
        self.pdf_entry = ttk.Entry(row, textvariable=self.pdf_var, state="readonly", cursor="hand2")
        self.pdf_entry.pack(side="left", fill="x", expand=True)
        self.pdf_entry.bind("<Button-1>", lambda e: self.choose_pdf())
        self.browse_btn = ttk.Button(row, text=self.t("browse"), command=self.choose_pdf)
        self.browse_btn.pack(side="left", padx=(8, 0))
        self.pdf_info = ttk.Label(c1, text="", style="Muted.TLabel")
        self.pdf_info.pack(anchor="w", pady=(8, 0))

        # Step 2 - Password
        _, c2 = self._card(root, "2", self.t("step2"))
        row = ttk.Frame(c2, style="Card.TFrame")
        row.pack(fill="x", pady=(10, 0))
        self.pw_entry = ttk.Entry(row, textvariable=self.pw_var,
                                  show="" if self.show_pw.get() else "•")
        self.pw_entry.pack(side="left", fill="x", expand=True)
        ttk.Checkbutton(row, text=self.t("show"), variable=self.show_pw,
                        style="Card.TCheckbutton",
                        command=self._toggle_show).pack(side="left", padx=(10, 0))
        self.pw_info = ttk.Label(c2, text="", style="Muted.TLabel")
        self.pw_info.pack(anchor="w", pady=(8, 0))

        # Step 3 - Output
        _, c3 = self._card(root, "3", self.t("step3"))
        row = ttk.Frame(c3, style="Card.TFrame")
        row.pack(fill="x", pady=(10, 0))
        self.out_entry = ttk.Entry(row, textvariable=self.out_var, state="readonly")
        self.out_entry.pack(side="left", fill="x", expand=True)
        self.out_btn = ttk.Button(row, text=self.t("change"), command=self.choose_output)
        self.out_btn.pack(side="left", padx=(8, 0))
        ttk.Label(c3, text=self.t("out_hint"), style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

        # Action + progress
        self.convert_btn = ttk.Button(root, style="Primary.TButton", command=self.start_convert)
        self.convert_btn.pack(fill="x", pady=(4, 10))
        self.progress = ttk.Progressbar(root, mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        self.status_label = ttk.Label(root, text="", style="Status.TLabel")
        self.status_label.pack(anchor="w", pady=(6, 0))

        # Result (hidden until done)
        self.result = tk.Frame(root, highlightthickness=1)
        inner = tk.Frame(self.result, padx=14, pady=12)
        inner.pack(fill="x")
        self.result_title = tk.Label(inner, font=(self.font, 11, "bold"), anchor="w",
                                     justify="left", wraplength=500)
        self.result_title.pack(fill="x")
        self.result_text = tk.Label(inner, fg=TEXT, font=(self.font, 9),
                                    anchor="w", justify="left", wraplength=500)
        self.result_text.pack(fill="x", pady=(2, 10))
        btns = tk.Frame(inner)
        btns.pack(fill="x")
        self.result_bg = [self.result, inner, self.result_title, self.result_text, btns]
        ttk.Button(btns, text=self.t("open_xlsx"), command=self.open_output).pack(side="left")
        ttk.Button(btns, text=self.t("show_folder"),
                   command=self.show_in_folder).pack(side="left", padx=8)
        ttk.Button(btns, text=self.t("another"), command=self.reset).pack(side="right")

        # Put the current state back on screen
        self._render_pdf_info()
        self._render_status()
        if self.last_result:
            self.progress["value"] = 100
            self._render_result()
        for entry, var in ((self.pdf_entry, self.pdf_var), (self.out_entry, self.out_var)):
            if var.get():  # new boxes need to be drawn first, so wait a little longer
                entry.after(250, lambda e=entry: e.xview_moveto(1))
        self._refresh()

    def set_lang(self, lang):
        if lang == self.lang or self.busy:
            return
        self.lang = lang
        save_lang(lang)
        self._build()

    # ---------- rendering of state ----------
    def _render_pdf_info(self):
        if self.pdf_error:
            self.pdf_info.configure(text=self.t("bad_pdf"), foreground=ERROR)
        elif self.pdf_path is None:
            self.pdf_info.configure(text=self.t("no_file"), foreground=MUTED)
        else:
            lock = self.t("locked") if self.locked else self.t("not_locked")
            self.pdf_info.configure(text=f"{human_size(self.pdf_path.stat().st_size)}  ·  {lock}",
                                    foreground=WARN if self.locked else MUTED)

    def _render_status(self):
        key, kw = self.status
        text = self.t(key, **kw) if key in STRINGS[self.lang] else key
        self.status_label.configure(text=text, foreground=self.status_color)

    def _set_status(self, key="", color=MUTED, **kw):
        self.status, self.status_color = (key, kw), color
        self._render_status()

    def _render_result(self):
        stats = self.last_result
        found = self.t("found", n=stats["transactions"], pages=stats["pages"])
        lines, warn = [], False
        if stats["format"] == "KBank":
            lines.append(self.t("account", acct=stats["account"] or self.t("acct_missing")))
            part = 0 if self.lang == "th" else -1  # check labels are "ไทย / English"
            failed = [label.split(" / ")[part] for label, ok in stats["checks"] if ok is False]
            unknown = all(ok is None for label, ok in stats["checks"][:5])
            if stats.get("skipped"):
                warn = True
                lines.append(self.t("skipped", n=stats["skipped"]))
            if failed:
                warn = True
                lines.append(self.t("mismatch", items=", ".join(failed)))
            elif unknown:
                lines.append(self.t("unknown"))
            else:
                lines.append(self.t("all_ok"))
        elif stats["transactions"] == 0:
            warn = True
            lines.append(self.t("none_found"))
        lines.append(self.t("saved", path=self.out_path))

        bg, border, fg = ("#FFF6DB", "#F0D48A", WARN) if warn else ("#E7F4EC", "#B7DFC5", SUCCESS)
        for w in self.result_bg:
            w.configure(bg=bg)
        self.result.configure(highlightbackground=border, highlightcolor=border)
        self.result_title.configure(fg=fg, text=self.t("done_check" if warn else "done",
                                                       found=found))
        self.result_text.configure(text="\n".join(lines))
        self.result.pack(fill="x", pady=(12, 0))

    def _refresh(self):
        """Enable/disable controls and set hints based on the current state."""
        has_pdf = self.pdf_path is not None
        idle = not self.busy

        self.pw_entry.configure(state="normal" if (has_pdf and self.locked and idle) else "disabled")
        if not has_pdf:
            self._pw_hint(self.t("pw_choose_first"), MUTED)
        elif not self.locked:
            self._pw_hint(self.t("pw_not_locked"), MUTED)
        elif self.pw_wrong:
            self._pw_hint(self.t("pw_wrong"), ERROR)
        elif self.pw_var.get():
            self._pw_hint(self.t("pw_ready"), MUTED)
        else:
            self._pw_hint(self.t("pw_enter"), WARN)

        self.browse_btn.configure(state="normal" if idle else "disabled")
        self.out_btn.configure(state="normal" if (has_pdf and idle) else "disabled")

        ready = has_pdf and idle and (not self.locked or self.pw_var.get())
        self.convert_btn.configure(state="normal" if ready else "disabled",
                                   text=self.t("converting") if self.busy else self.t("convert"))

    def _pw_hint(self, text, color):
        self.pw_info.configure(text=text, foreground=color)

    def _on_password_typed(self):
        self.pw_wrong = False
        if hasattr(self, "pw_entry"):
            self._refresh()

    def _toggle_show(self):
        self.pw_entry.configure(show="" if self.show_pw.get() else "•")

    # ---------- actions ----------
    def choose_pdf(self):
        if self.busy:
            return
        start = Path(self.pdf_path).parent if self.pdf_path else Path.home() / "Downloads"
        path = filedialog.askopenfilename(
            parent=self, title=self.t("pick_title"), initialdir=start,
            filetypes=[(self.t("pdf_files"), "*.pdf"), (self.t("all_files"), "*.*")])
        if not path:
            return
        path = Path(path)

        self.last_result = None
        self.result.pack_forget()
        self.progress["value"] = 0
        self.pw_var.set("")
        self._set_status("")
        try:
            self.locked = is_locked(path)
            self.pdf_error = False
        except Exception:
            self.pdf_path = self.out_path = None
            self.pdf_error = True
            self.pdf_var.set("")
            self.out_var.set("")
            self._render_pdf_info()
            self._refresh()
            return

        self.pdf_path = path
        self._show_path(self.pdf_entry, self.pdf_var, path)
        self.out_path = path.with_suffix(".xlsx")
        self._show_path(self.out_entry, self.out_var, self.out_path)
        self._render_pdf_info()
        self._refresh()
        (self.pw_entry if self.locked else self.convert_btn).focus_set()

    def choose_output(self):
        if not self.pdf_path or self.busy:
            return
        path = filedialog.asksaveasfilename(
            parent=self, title=self.t("save_title"), defaultextension=".xlsx",
            initialdir=self.out_path.parent, initialfile=self.out_path.name,
            filetypes=[(self.t("excel_files"), "*.xlsx")])
        if path:
            self.out_path = Path(path)
            self._show_path(self.out_entry, self.out_var, self.out_path)

    def _show_path(self, entry, var, path):
        """Show a path, scrolled so the file name at the end is visible."""
        var.set(str(path))
        entry.after(50, lambda: entry.xview_moveto(1))

    def start_convert(self):
        if str(self.convert_btn["state"]) == "disabled":
            return
        if self.out_path.exists() and not messagebox.askyesno(
                self.t("replace_title"), self.t("replace_msg", name=self.out_path.name),
                parent=self):
            return

        self.busy = True
        self.last_result = None
        self.result.pack_forget()
        self.progress["value"] = 0
        self._set_status("opening")
        self._refresh()

        args = (self.pdf_path, self.out_path, self.pw_var.get() or None)
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, pdf, out, password):
        try:
            stats = convert(pdf, out, password,
                            progress=lambda d, t: self.events.put(("progress", d, t)))
            self.events.put(("done", stats))
        except PasswordError:
            self.events.put(("password",))
        except PermissionError:
            self.events.put(("error", "err_save", {}))
        except Exception as exc:
            self.events.put(("error", "err_other", {"exc": exc}))

    def _poll_events(self):
        try:
            while True:
                self._handle(self.events.get_nowait())
        except queue.Empty:
            pass
        self.after(50, self._poll_events)

    def _handle(self, ev):
        kind = ev[0]
        if kind == "progress":
            _, done, total = ev
            self.progress["value"] = done / total * 100
            self._set_status("reading", done=done, total=total)
            return

        self.busy = False
        if kind == "done":
            self.last_result = ev[1]
            self.pw_var.set("")  # don't keep the password around
            self.progress["value"] = 100
            self._set_status("")
            self._render_result()
        elif kind == "password":
            self.progress["value"] = 0
            self._set_status("")
            self.pw_wrong = True
            self._refresh()
            self.pw_entry.focus_set()
            self.pw_entry.select_range(0, "end")
            return
        else:
            self.progress["value"] = 0
            self._set_status(ev[1], ERROR, **ev[2])
        self._refresh()

    def open_output(self):
        if self.out_path and self.out_path.exists():
            os.startfile(self.out_path)

    def show_in_folder(self):
        if self.out_path and self.out_path.exists():
            subprocess.run(["explorer", "/select,", str(self.out_path)])

    def reset(self):
        self.pdf_path = self.out_path = self.last_result = None
        self.locked = self.pdf_error = self.pw_wrong = False
        self.pdf_var.set("")
        self.out_var.set("")
        self.pw_var.set("")
        self.progress["value"] = 0
        self._set_status("")
        self.result.pack_forget()
        self._render_pdf_info()
        self._refresh()


def resource(name):
    """Path to a file shipped with the app (works both as .py and inside the .exe)."""
    base = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return Path(base) / name


def main():
    # Hidden test mode: app.exe --convert in.pdf out.xlsx (password in STE_PASSWORD).
    # Exit code 0 = converted, 1 = failed.
    if len(sys.argv) == 4 and sys.argv[1] == "--convert":
        try:
            convert(sys.argv[2], sys.argv[3], os.environ.get("STE_PASSWORD") or None)
            sys.exit(0)
        except Exception:
            sys.exit(1)

    if sys.platform == "win32":
        try:  # sharp text on high-DPI screens
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App()
    try:
        app.iconbitmap(default=str(resource("icon.ico")))
    except tk.TclError:
        pass
    app.mainloop()


if __name__ == "__main__":
    main()
