#!/usr/bin/env python3
"""
DocDivide -- Engineering Drawing PDF Splitter
---------------------------------------------
Splits a multi-drawing PDF into individual files by drawing number.
Extracts title block data (drawing number, revision, description, sheet)
using the Anthropic Claude API (vision). Auto-deduplicates sheets.

Requirements:
    pip install anthropic pypdf2 pdf2image pillow ttkbootstrap

Also requires poppler for pdf2image:
    Windows: https://github.com/oschwartz10612/poppler-windows/releases
             Extract and add bin/ folder to PATH
    Mac:     brew install poppler
    Linux:   sudo apt install poppler-utils
"""

import os
import sys
import json
import base64
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from io import BytesIO
from pathlib import Path

# -- Dependency check --
missing = []
try:
    import anthropic
except ImportError:
    missing.append("anthropic")
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        missing.append("pypdf2")
try:
    from pdf2image import convert_from_path
except ImportError:
    missing.append("pdf2image")
try:
    from PIL import Image
except ImportError:
    missing.append("pillow")

if missing:
    print(f"Missing packages: {', '.join(missing)}")
    print(f"Run: pip install {' '.join(missing)}")
    sys.exit(1)

try:
    import ttkbootstrap as ttk_bs
    from ttkbootstrap.constants import *
    USE_BOOTSTRAP = True
except ImportError:
    USE_BOOTSTRAP = False

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    USE_DND = True
except ImportError:
    USE_DND = False

import csv
import re
import zipfile

# -- Embedded API key (obfuscated with XOR) --
# Run embed_key.py to regenerate these values for your key.
_SALT = b'\xc0\xafI\x80W\x12\xbb&\xab\x12\x86\x8bq(\xaa\xa9\x89Mp\x82>]g\xbd\xcfZd\x9d\xf3:\xf9\xf7'
_ENCODED_KEY = b'\xb3\xc4d\xe19f\x96G\xdb{\xb6\xb8\\N\xcd\xe0\xc7\x12\x07\xeal0,\xfb\xfb\x10T\xd4\x98J\xab\x8d\xa3\xed\x08\xe6\x16_\xdaU\xf4D\xb2\xbeCd\xc2\x84\xbc/5\xcfJ\x15?\xec\xb8m)\xec\xb0\x0c\xb4\x9a\xf3\xe6\x0e\xe9ft\xf2G\xec[\xb1\xcf&p\x9b\x98\xeb#!\xceZ%\x08\xc7\xfe\x1f\x0c\xab\xb1l\xcd\xa8\x86\xf8\x18\xad\x06j\xd4H\xc5u\xc7\xca'


def _get_embedded_key() -> str:
    if not _SALT or not _ENCODED_KEY:
        return ""
    salt = _SALT * (len(_ENCODED_KEY) // len(_SALT) + 1)
    return bytes(a ^ b for a, b in zip(_ENCODED_KEY, salt)).decode()


MODEL = "claude-sonnet-4-20250514"
ERP_HOST = "PK8"
ERP_PORT = 1521
ERP_SERVICE = "LIVE1"
ERP_USER = "PK1"
ERP_PWD = "PK1"

# -- Helpers --

def page_to_base64(pdf_path: str, page_idx: int) -> str:
    images = convert_from_path(pdf_path, first_page=page_idx + 1, last_page=page_idx + 1, dpi=150)
    img = images[0]
    w, h = img.size
    tb_w = int(w * 0.60)
    tb_h = int(h * 0.32)
    crop = img.crop((w - tb_w, h - tb_h, w, h))
    buf = BytesIO()
    crop.save(buf, format="JPEG", quality=88)
    return base64.b64encode(buf.getvalue()).decode()


def extract_title_block(client, b64_img: str, page_num: int) -> dict:
    prompt = f"""You are reading an engineering drawing title block. PDF page: {page_num + 1}.

CRITICAL -- drawing_number rules:
- Short alphanumeric code: "A-101", "M-203", "DWG-0042", "C100", "E-3.1"
- Found in a box labeled "DWG NO", "DRAWING NO", "DRAWING NUMBER", "DOC NO", etc.
- NOT the project name, client name, contract, or description
- Under 20 characters; if it reads like a sentence it is wrong
- If you see project number AND drawing number, use only the drawing number

Extract:
- drawing_number: unique sheet identifier
- revision: revision/edition label (REV, REVISION, EDITION). null if absent.
- description: drawing title/description
- sheet: sheet indicator "1/3", "2 OF 5", etc. null if single-sheet.
- project: project name/number (for context only)

Return ONLY valid JSON, no markdown:
{{"drawing_number":"...","revision":"...","description":"...","sheet":"...","project":"..."}}
Use null for missing fields."""

    msg = client.messages.create(
        model=MODEL,
        max_tokens=500,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64_img}},
                {"type": "text", "text": prompt}
            ]
        }]
    )
    text = msg.content[0].text.strip()
    text = re.sub(r"```json|```", "", text).strip()
    try:
        return json.loads(text)
    except Exception:
        return {"drawing_number": None, "revision": None, "description": None, "sheet": None, "project": None}


def parse_sheet(sheet):
    if not sheet:
        return None, None
    m = re.search(r"(\d+)\s*[/\sOFof]+\s*(\d+)", str(sheet), re.IGNORECASE)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def is_suspect(dn, description, project):
    if not dn:
        return True
    if len(dn) > 30:
        return True
    if len(dn.strip().split()) > 5:
        return True
    if description and dn.lower() == description.lower():
        return True
    if project and dn.lower() == project.lower():
        return True
    return False


def deduplicate_pages(pages):
    by_sheet = {}
    for p in pages:
        key = str(p["sheet_current"]) if p["sheet_current"] is not None else f"_pg{p.get('pdf_idx', 0)}_{p['page']}"
        by_sheet[key] = p
    kept_ids = {(p.get("pdf_path", ""), p["page"]) for p in by_sheet.values()}
    kept = [p for p in pages if (p.get("pdf_path", ""), p["page"]) in kept_ids]
    removed = [p for p in pages if (p.get("pdf_path", ""), p["page"]) not in kept_ids]
    return kept, removed


# -- GUI --

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("DocDivide -- Engineering Drawing PDF Splitter")
        self.root.geometry("1000x720")
        self.root.resizable(True, True)

        self.pdf_paths = []
        self.drawings = []
        self.removed_log = []
        self.cancel_flag = threading.Event()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        top = tk.Frame(self.root, bg="#1e3a5f", height=56)
        top.pack(fill=tk.X)
        tk.Label(top, text="DocDivide  --  Engineering Drawing PDF Splitter",
                 bg="#1e3a5f", fg="white", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT, padx=16, pady=12)

        main = tk.Frame(self.root, bg="#f0f4f8")
        main.pack(fill=tk.BOTH, expand=True, padx=14, pady=10)

        row1 = tk.LabelFrame(main, text="Setup", bg="#f0f4f8", font=("Segoe UI", 10, "bold"))
        row1.pack(fill=tk.X, pady=(0, 8))

        tk.Label(row1, text="PDF File(s):", bg="#f0f4f8").grid(row=0, column=0, sticky="w", **pad)
        self.file_var = tk.StringVar(value="No file selected")
        tk.Label(row1, textvariable=self.file_var, bg="#f0f4f8", fg="#334155", width=50, anchor="w").grid(row=0, column=1, sticky="w", **pad)
        tk.Button(row1, text="Browse...", command=self._browse_files, bg="#e2e8f0").grid(row=0, column=2, **pad)

        row2 = tk.Frame(main, bg="#f0f4f8")
        row2.pack(fill=tk.X, pady=(0, 8))

        self.scan_btn = tk.Button(row2, text="Start Scanning", command=self._start_scan,
                                   bg="#1e3a5f", fg="white", font=("Segoe UI", 11, "bold"),
                                   padx=14, pady=6, relief=tk.FLAT, cursor="hand2")
        self.scan_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.cancel_btn = tk.Button(row2, text="Cancel", command=self._cancel,
                                     bg="#dc2626", fg="white", padx=10, pady=6, relief=tk.FLAT,
                                     cursor="hand2", state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT)

        self.status_var = tk.StringVar(value="Upload a PDF to begin.")
        tk.Label(row2, textvariable=self.status_var, bg="#f0f4f8", fg="#475569",
                 wraplength=600, justify=tk.LEFT).pack(side=tk.LEFT, padx=14)

        self.progress = ttk.Progressbar(main, mode="determinate", length=200)
        self.progress.pack(fill=tk.X, pady=(0, 6))

        table_frame = tk.LabelFrame(main, text="Drawing Index  (double-click a cell to edit)",
                                    bg="#f0f4f8", font=("Segoe UI", 10, "bold"))
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        cols = ("#", "Drawing Number", "Revision", "Description", "Sheets", "PDF Pages", "Flags", "ERP Rev", "ERP Status")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=14)
        widths = [36, 140, 70, 300, 56, 120, 80, 70, 90]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=w)
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.tag_configure("suspect", background="#fef9c3")
        self.tree.tag_configure("normal", background="white")
        self.tree.tag_configure("erp_match", background="#dcfce7")
        self.tree.tag_configure("erp_mismatch", background="#fef9c3")
        self.tree.tag_configure("erp_missing", background="#fee2e2")
        self.tree.bind("<Double-1>", self._on_double_click)

        if USE_DND:
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind("<<Drop>>", self._on_drop)

        row4 = tk.Frame(main, bg="#f0f4f8")
        row4.pack(fill=tk.X)

        self.split_btn = tk.Button(row4, text="Split & Save ZIP", command=self._split_and_save,
                                    bg="#16a34a", fg="white", font=("Segoe UI", 11, "bold"),
                                    padx=14, pady=6, relief=tk.FLAT, cursor="hand2", state=tk.DISABLED)
        self.split_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.csv_btn = tk.Button(row4, text="Export CSV Only", command=self._export_csv,
                                  bg="#475569", fg="white", padx=10, pady=6, relief=tk.FLAT,
                                  cursor="hand2", state=tk.DISABLED)
        self.csv_btn.pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(row4, text="Merge Selected Up", command=self._merge_up,
                  bg="#e0f2fe", relief=tk.FLAT, padx=8, pady=6).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(row4, text="Split Selected", command=self._split_selected,
                  bg="#fce7f3", relief=tk.FLAT, padx=8, pady=6).pack(side=tk.LEFT)

        self.removed_var = tk.StringVar(value="")
        tk.Label(row4, textvariable=self.removed_var, bg="#f0f4f8",
                 fg="#7e22ce", font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=10)

    def _browse_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.pdf_paths = list(paths)
            count = len(paths)
            label = Path(paths[0]).name if count == 1 else f"{count} files selected"
            self.file_var.set(label)

    def _on_drop(self, event):
        paths = [p for p in self.tk.splitlist(event.data) if p.lower().endswith(".pdf")]
        if not paths:
            return
        self.pdf_paths = paths
        count = len(paths)
        label = Path(paths[0]).name if count == 1 else f"{count} files selected"
        self.file_var.set(label)
        self.status_var.set(f"{count} PDF(s) loaded. Click Start Scanning.")

    def _cancel(self):
        self.cancel_flag.set()
        self.status_var.set("Cancelling...")

    def _set_ui_scanning(self, scanning: bool):
        state = tk.DISABLED if scanning else tk.NORMAL
        self.scan_btn.config(state=state)
        self.cancel_btn.config(state=tk.NORMAL if scanning else tk.DISABLED)

    def _start_scan(self):
        if not self.pdf_paths:
            messagebox.showwarning("No File", "Please select at least one PDF file.")
            return
        key = _get_embedded_key()
        if not key:
            messagebox.showerror("API Key Missing", "No API key is embedded.\nRun embed_key.py and rebuild the application.")
            return
        self.cancel_flag.clear()
        self.drawings = []
        self.removed_log = []
        self._refresh_table()
        self.split_btn.config(state=tk.DISABLED)
        self.csv_btn.config(state=tk.DISABLED)
        self._set_ui_scanning(True)
        threading.Thread(target=self._scan_thread, args=(key,), daemon=True).start()

    def _scan_thread(self, api_key: str):
        client = anthropic.Anthropic(api_key=api_key)
        try:
            readers = [PdfReader(p) for p in self.pdf_paths]
            total = sum(len(r.pages) for r in readers)
            self._update_status(f"Scanning {total} pages across {len(self.pdf_paths)} file(s)...")

            results = []
            done = 0
            for j, (pdf_path, reader) in enumerate(zip(self.pdf_paths, readers)):
                for i in range(len(reader.pages)):
                    if self.cancel_flag.is_set():
                        break
                    self._update_status(f"Scanning page {done + 1} of {total}...")
                    self._update_progress(int(done / total * 100))
                    try:
                        b64 = page_to_base64(pdf_path, i)
                        data = extract_title_block(client, b64, i)
                        results.append({"page": i, "pdf_path": pdf_path, "pdf_idx": j, **data})
                    except Exception as e:
                        results.append({"page": i, "pdf_path": pdf_path, "pdf_idx": j,
                                        "drawing_number": None, "revision": None,
                                        "description": None, "sheet": None, "project": None})
                    done += 1
                if self.cancel_flag.is_set():
                    break

            groups = {}
            for r in results:
                dn = r["drawing_number"] or f"UNKNOWN_PAGE_{r['page'] + 1}"
                if dn not in groups:
                    groups[dn] = {"drawing_number": dn, "revision": r["revision"],
                                  "description": r["description"], "project": r["project"],
                                  "erp_rev": "", "erp_status": "", "pages": []}
                sc, _ = parse_sheet(r["sheet"])
                groups[dn]["pages"].append({"page": r["page"], "pdf_path": r["pdf_path"],
                                            "pdf_idx": r["pdf_idx"], "sheet_current": sc, "sheet": r["sheet"]})

            for g in groups.values():
                g["pages"].sort(key=lambda p: (p["sheet_current"] or 9999, p["page"]))

            drawing_list = list(groups.values())
            removed_log = []
            for idx, d in enumerate(drawing_list):
                kept, removed = deduplicate_pages(d["pages"])
                if removed:
                    removed_log.append({"drawing_idx": idx, "drawing_number": d["drawing_number"], "removed_pages": removed})
                    d["pages"] = kept

            self.drawings = drawing_list
            self.removed_log = removed_log
            total_removed = sum(len(e["removed_pages"]) for e in removed_log)
            suspect_count = sum(1 for d in drawing_list if is_suspect(d["drawing_number"], d.get("description"), d.get("project")))

            erp_msg = self._check_erp()

            self.root.after(0, self._refresh_table)
            msg = f"Found {len(drawing_list)} drawings."
            if suspect_count:
                msg += f"  {suspect_count} suspect drawing numbers (highlighted)."
            if total_removed:
                msg += f"  {total_removed} duplicate pages auto-removed."
            if erp_msg:
                msg += f"  {erp_msg}"
            self._update_status(msg)
            self._update_progress(100)
            if total_removed:
                self.root.after(0, lambda: self.removed_var.set(
                    f"{total_removed} duplicate page(s) auto-removed across {len(removed_log)} drawing(s)"))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self._update_status(f"Error: {e}")
        finally:
            self.root.after(0, lambda: self._set_ui_scanning(False))
            self.root.after(0, lambda: self.split_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.csv_btn.config(state=tk.NORMAL))

    def _refresh_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for i, d in enumerate(self.drawings):
            suspect = is_suspect(d["drawing_number"], d.get("description"), d.get("project"))
            if len(self.pdf_paths) > 1:
                pages_str = ", ".join(f"F{p['pdf_idx']+1}:p{p['page']+1}" for p in d["pages"])
            else:
                pages_str = ", ".join(str(p["page"] + 1) for p in d["pages"])
            flags = "suspect" if suspect else ""
            erp_rev = d.get("erp_rev", "")
            erp_status = d.get("erp_status", "")
            if suspect:
                tag = "suspect"
            elif erp_status == "Match":
                tag = "erp_match"
            elif erp_status == "Mismatch":
                tag = "erp_mismatch"
            elif erp_status == "Not Found":
                tag = "erp_missing"
            else:
                tag = "normal"
            self.tree.insert("", tk.END, iid=str(i), tags=(tag,), values=(
                i + 1,
                d["drawing_number"] or "",
                d["revision"] or "",
                d["description"] or "",
                len(d["pages"]),
                pages_str,
                flags,
                erp_rev,
                erp_status
            ))

    def _on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        col_idx = int(col.replace("#", "")) - 1
        if col_idx not in (1, 2, 3):
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        field_map = {1: "drawing_number", 2: "revision", 3: "description"}
        field = field_map[col_idx]
        current_val = self.drawings[int(iid)][field] or ""
        x, y, w, h = self.tree.bbox(iid, col)
        popup = tk.Entry(self.tree, font=("Segoe UI", 10))
        popup.place(x=x, y=y, width=max(w, 180), height=h)
        popup.insert(0, current_val)
        popup.focus_set()

        def save(event=None):
            self.drawings[int(iid)][field] = popup.get()
            popup.destroy()
            self._refresh_table()

        popup.bind("<Return>", save)
        popup.bind("<FocusOut>", save)

    def _get_selected_idx(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No Selection", "Please select a row in the table.")
            return None
        return int(sel[0])

    def _merge_up(self):
        idx = self._get_selected_idx()
        if idx is None or idx == 0:
            messagebox.showinfo("Cannot Merge", "Select a row below the first row to merge up.")
            return
        target = self.drawings[idx - 1]
        src = self.drawings[idx]
        combined = target["pages"] + src["pages"]
        combined.sort(key=lambda p: (p["sheet_current"] or 9999, p["page"]))
        kept, removed = deduplicate_pages(combined)
        target["pages"] = kept
        if removed:
            self.removed_log.append({"drawing_idx": idx - 1, "drawing_number": target["drawing_number"], "removed_pages": removed})
        del self.drawings[idx]
        self._refresh_table()

    def _split_selected(self):
        idx = self._get_selected_idx()
        if idx is None:
            return
        d = self.drawings[idx]
        if len(d["pages"]) <= 1:
            messagebox.showinfo("Cannot Split", "This drawing only has one page.")
            return
        new_rows = []
        for pi, p in enumerate(d["pages"]):
            new_rows.append({**d, "drawing_number": d["drawing_number"] + (f"_SH{pi + 1}" if pi > 0 else ""), "pages": [p]})
        self.drawings[idx:idx + 1] = new_rows
        self._refresh_table()

    def _split_and_save(self):
        if not self.drawings:
            return
        out_path = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("ZIP archive", "*.zip")],
            title="Save ZIP as...",
            initialfile="engineering_drawings.zip"
        )
        if not out_path:
            return
        threading.Thread(target=self._save_thread, args=(out_path,), daemon=True).start()

    def _save_thread(self, zip_path: str):
        self._update_status("Building PDFs...")
        self.root.after(0, lambda: self.split_btn.config(state=tk.DISABLED))
        try:
            readers = {}

            def get_reader(path):
                if path not in readers:
                    readers[path] = PdfReader(path)
                return readers[path]

            total = len(self.drawings)
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, d in enumerate(self.drawings):
                    self._update_status(f"Writing {i + 1}/{total}: {d['drawing_number']}")
                    self._update_progress(int(i / total * 100))
                    writer = PdfWriter()
                    for p in d["pages"]:
                        writer.add_page(get_reader(p["pdf_path"]).pages[p["page"]])
                    safe_name = re.sub(r"[^\w\-\.]", "_", str(d["drawing_number"]))
                    buf = BytesIO()
                    writer.write(buf)
                    zf.writestr(f"{safe_name}.pdf", buf.getvalue())

                csv_text = "Drawing Number,Revision,Description,Sheet Count,ERP Rev,ERP Status\n"
                for d in self.drawings:
                    desc = (d.get("description") or "").replace('"', '""')
                    csv_text += f'"{d["drawing_number"] or ""}","{d.get("revision") or ""}","{desc}","{len(d["pages"])}","{d.get("erp_rev") or ""}","{d.get("erp_status") or ""}"\n'
                zf.writestr("drawing_index.csv", csv_text)

            self._update_status(f"Done! Saved {total} drawings to {Path(zip_path).name}")
            self._update_progress(100)
            self.root.after(0, lambda: messagebox.showinfo("Complete", f"Saved {total} PDFs + CSV to:\n{zip_path}"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self._update_status(f"Error: {e}")
        finally:
            self.root.after(0, lambda: self.split_btn.config(state=tk.NORMAL))

    def _export_csv(self):
        if not self.drawings:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV", "*.csv")],
            title="Save CSV as...", initialfile="drawing_index.csv")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Drawing Number", "Revision", "Description", "Sheet Count", "ERP Rev", "ERP Status"])
            for d in self.drawings:
                writer.writerow([d["drawing_number"] or "", d.get("revision") or "",
                                  d.get("description") or "", len(d["pages"]),
                                  d.get("erp_rev") or "", d.get("erp_status") or ""])
        messagebox.showinfo("Saved", f"CSV saved to {path}")

    def _check_erp(self) -> str:
        """Query ProfitKey ERP for rev levels. Returns a status note or empty string on failure."""
        part_numbers = [d["drawing_number"] for d in self.drawings if d["drawing_number"]]
        if not part_numbers:
            return ""
        try:
            import oracledb
            placeholders = ",".join(f":{i+1}" for i in range(len(part_numbers)))
            sql = f"SELECT DISTINCT IM_KEY, IM_REV FROM PK1.IM WHERE IM_KEY IN ({placeholders})"
            conn = oracledb.connect(user=ERP_USER, password=ERP_PWD,
                                    dsn=f"{ERP_HOST}:{ERP_PORT}/{ERP_SERVICE}")
            cursor = conn.cursor()
            cursor.execute(sql, part_numbers)
            erp_data = {row[0].strip(): row[1].strip() if row[1] else "" for row in cursor}
            conn.close()
            for d in self.drawings:
                dn = d["drawing_number"]
                erp_rev = erp_data.get(dn)
                if erp_rev is None:
                    d["erp_rev"] = ""
                    d["erp_status"] = "Not Found"
                elif erp_rev.upper() == (d.get("revision") or "").strip().upper():
                    d["erp_rev"] = erp_rev
                    d["erp_status"] = "Match"
                else:
                    d["erp_rev"] = erp_rev
                    d["erp_status"] = "Mismatch"
            match = sum(1 for d in self.drawings if d["erp_status"] == "Match")
            mismatch = sum(1 for d in self.drawings if d["erp_status"] == "Mismatch")
            missing = sum(1 for d in self.drawings if d["erp_status"] == "Not Found")
            return f"ERP: {match} match, {mismatch} mismatch, {missing} not found."
        except Exception as e:
            return f"ERP check unavailable: {e}"

    def _update_status(self, msg: str):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _update_progress(self, val: int):
        self.root.after(0, lambda: self.progress.config(value=val))


if __name__ == "__main__":
    root = TkinterDnD.Tk() if USE_DND else tk.Tk()
    app = App(root)
    root.mainloop()
