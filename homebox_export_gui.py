#!/usr/bin/env python3
"""
Homebox Export Tool — GUI Application
Generates professional PDF inventory reports from your Homebox instance.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import queue
import os
import sys
import io
import json
from datetime import datetime
from pathlib import Path

import requests
from fpdf import FPDF
from PIL import Image


# ═══════════════════════════════════════════════════════════════════════════════
# Color Palette (PDF)
# ═══════════════════════════════════════════════════════════════════════════════
NAVY       = (26, 54, 93)
SLATE      = (51, 65, 85)
GRAY_700   = (55, 65, 81)
GRAY_500   = (107, 114, 128)
GRAY_200   = (229, 231, 235)
GRAY_100   = (243, 244, 246)
GRAY_50    = (249, 250, 251)
WHITE      = (255, 255, 255)
BLUE_600   = (37, 99, 235)

def _get_config_path() -> Path:
    """Config lives next to the .exe (or .py script) so it's portable."""
    if getattr(sys, "frozen", False):
        # Running as PyInstaller .exe
        return Path(sys.executable).parent / "homebox_export_config.json"
    else:
        # Running as .py script
        return Path(__file__).parent / "homebox_export_config.json"


# ═══════════════════════════════════════════════════════════════════════════════
# Homebox API Client
# ═══════════════════════════════════════════════════════════════════════════════
class HomeboxClient:
    def __init__(self, base_url: str):
        self.base_url = base_url.rstrip("/")
        self.api = f"{self.base_url}/api/v1"
        self.session = requests.Session()
        self.session.headers["Accept"] = "application/json"

    def login(self, username: str, password: str) -> dict:
        r = self.session.post(f"{self.api}/users/login", json={
            "username": username, "password": password,
        })
        r.raise_for_status()
        data = r.json()
        token = data.get("token", "")
        if token.startswith("Bearer "):
            self.session.headers["Authorization"] = token
        else:
            self.session.headers["Authorization"] = f"Bearer {token}"
        att_token = data.get("attachmentToken", "")
        if att_token:
            self.session.cookies.set("hb.auth.attachment_token", att_token)
        return data

    def get_items(self, query: str = None) -> list:
        params = {"pageSize": 100000}
        if query:
            params["q"] = query
        r = self.session.get(f"{self.api}/items", params=params)
        r.raise_for_status()
        data = r.json()
        if isinstance(data, list):
            return data
        return data.get("items", [])

    def get_item(self, item_id: str) -> dict:
        r = self.session.get(f"{self.api}/items/{item_id}")
        r.raise_for_status()
        return r.json()

    def get_maintenance(self, item_id: str) -> list:
        r = self.session.get(f"{self.api}/items/{item_id}/maintenance")
        r.raise_for_status()
        data = r.json()
        if isinstance(data, list):
            return data
        return data.get("entries", data.get("data", []))

    def get_attachment(self, item_id: str, attachment_id: str) -> bytes:
        r = self.session.get(
            f"{self.api}/items/{item_id}/attachments/{attachment_id}",
            stream=True,
        )
        r.raise_for_status()
        return r.content

    @staticmethod
    def format_asset_id(raw) -> str:
        try:
            n = int(str(raw).replace("-", ""))
        except (ValueError, TypeError):
            n = 0
        s = str(n).zfill(9)
        return f"{s[:3]}-{s[3:6]}-{s[6:]}"


# ═══════════════════════════════════════════════════════════════════════════════
# PDF Report Generator
# ═══════════════════════════════════════════════════════════════════════════════
class InventoryReport(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="Letter")
        self.set_auto_page_break(True, margin=25)
        self.set_margins(20, 20, 20)
        self._pw = self.w - 40
        self.owner = ""
        self.total_value = 0.0
        self.item_count = 0
        self._is_cover = True

    @staticmethod
    def _t(text) -> str:
        if text is None:
            return ""
        text = str(text)
        for orig, repl in {
            "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
            "\u2013": "-", "\u2014": "--", "\u2026": "...", "\u2022": "*",
            "\u00a0": " ",
        }.items():
            text = text.replace(orig, repl)
        return text.encode("latin-1", errors="replace").decode("latin-1")

    @staticmethod
    def _date(val) -> str:
        if not val:
            return ""
        s = str(val)
        if s.startswith("0001"):
            return ""
        try:
            if "T" in s:
                dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
            else:
                dt = datetime.strptime(s[:10], "%Y-%m-%d")
            return dt.strftime("%b %d, %Y")
        except (ValueError, TypeError):
            return s

    @staticmethod
    def _price(val) -> str:
        try:
            f = float(val or 0)
            return f"${f:,.2f}" if f else ""
        except (ValueError, TypeError):
            return ""

    @staticmethod
    def _loc_name(item) -> str:
        loc = item.get("location")
        return loc.get("name", "") if isinstance(loc, dict) else ""

    def header(self):
        if self._is_cover:
            return
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*GRAY_500)
        self.cell(self._pw / 2, 8, "Home Inventory Report")
        self.cell(self._pw / 2, 8, datetime.now().strftime("%B %d, %Y"), align="R")
        self.ln()
        self.set_draw_color(*GRAY_200)
        self.set_line_width(0.3)
        self.line(20, self.get_y(), self.w - 20, self.get_y())
        self.ln(4)

    def footer(self):
        if self._is_cover:
            return
        self.set_y(-18)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*GRAY_500)
        self.cell(0, 8, f"Page {self.page_no() - 1}", align="C")

    def add_cover(self, owner: str = ""):
        self.owner = owner
        self._is_cover = True
        self.add_page()
        self.set_fill_color(*NAVY)
        self.rect(0, 0, self.w, 100, "F")
        self.set_y(28)
        self.set_font("Helvetica", "B", 30)
        self.set_text_color(*WHITE)
        self.cell(0, 14, "HOME INVENTORY", align="C")
        self.ln()
        self.set_font("Helvetica", "", 22)
        self.cell(0, 11, "REPORT", align="C")
        self.ln(8)
        self.set_draw_color(*WHITE)
        self.set_line_width(0.4)
        cx = self.w / 2
        self.line(cx - 25, self.get_y(), cx + 25, self.get_y())
        self.ln(7)
        self.set_font("Helvetica", "I", 11)
        self.cell(0, 7, "For Insurance & Documentation Purposes", align="C")
        self.set_y(115)
        self.set_text_color(*GRAY_700)
        self.set_font("Helvetica", "", 12)
        if owner:
            self.cell(0, 9, f"Prepared for:  {self._t(owner)}", align="C")
            self.ln()
        self.cell(0, 9, f"Date:  {datetime.now().strftime('%B %d, %Y')}", align="C")
        self.ln(20)
        self.set_font("Helvetica", "I", 9)
        self.set_text_color(*GRAY_500)
        self.cell(0, 6, "Generated from Homebox Inventory Management System", align="C")
        self._is_cover = False

    def add_summary(self, items: list):
        self.add_page()
        self.set_font("Helvetica", "B", 18)
        self.set_text_color(*NAVY)
        self.cell(0, 10, "Inventory Summary")
        self.ln(14)
        total_val = sum(float(it.get("purchasePrice") or 0) for it in items)
        insured = sum(1 for it in items if it.get("insured"))
        stats = [
            ("Total Items", str(len(items))),
            ("Estimated Value", f"${total_val:,.2f}"),
            ("Insured Items", str(insured)),
        ]
        box_w = self._pw / 3 - 3
        y = self.get_y()
        for i, (label, value) in enumerate(stats):
            x = 20 + i * (box_w + 4.5)
            self.set_fill_color(*GRAY_100)
            self.rect(x, y, box_w, 22, "F")
            self.set_xy(x, y + 3)
            self.set_font("Helvetica", "", 8)
            self.set_text_color(*GRAY_500)
            self.cell(box_w, 5, label, align="C")
            self.set_xy(x, y + 10)
            self.set_font("Helvetica", "B", 14)
            self.set_text_color(*NAVY)
            self.cell(box_w, 7, value, align="C")
        self.set_y(y + 30)

        cols = [26, 60, 38, 24, 22]
        headers = ["Asset ID", "Name", "Location", "Value", "Insured"]
        self._table_hdr(cols, headers)
        self.set_font("Helvetica", "", 8)
        for idx, item in enumerate(items):
            if self.get_y() > self.h - 30:
                self.add_page()
                self._table_hdr(cols, headers)
                self.set_font("Helvetica", "", 8)
            bg = GRAY_50 if idx % 2 == 0 else WHITE
            self.set_fill_color(*bg)
            self.set_text_color(*GRAY_700)
            aid = HomeboxClient.format_asset_id(item.get("assetId", 0))
            self.cell(cols[0], 6, f" {aid}", fill=True)
            self.cell(cols[1], 6, f" {self._t(item.get('name', ''))[:33]}", fill=True)
            self.cell(cols[2], 6, f" {self._t(self._loc_name(item))[:20]}", fill=True)
            self.cell(cols[3], 6, f"{self._price(item.get('purchasePrice'))} ", fill=True, align="R")
            self.cell(cols[4], 6, "Yes" if item.get("insured") else "", fill=True, align="C")
            self.ln()

    def _table_hdr(self, cols, headers):
        self.set_fill_color(*NAVY)
        self.set_text_color(*WHITE)
        self.set_font("Helvetica", "B", 8)
        for i, h in enumerate(headers):
            al = "R" if h == "Value" else ("C" if h == "Insured" else "L")
            self.cell(cols[i], 7, f" {h}", fill=True, align=al)
        self.ln()

    def _heading(self, title):
        if self.get_y() > self.h - 28:
            self.add_page()
        self.ln(2)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*NAVY)
        self.cell(0, 6, self._t(title))
        self.ln()
        self.set_draw_color(*BLUE_600)
        self.set_line_width(0.3)
        self.line(20, self.get_y(), self.w - 20, self.get_y())
        self.ln(3)

    def _row(self, label, value, label_w=38):
        if not value:
            return
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*GRAY_500)
        self.cell(label_w, 5, self._t(label))
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*GRAY_700)
        val = self._t(str(value))
        if len(val) > 65:
            self.ln()
            self.set_x(20)
            self.multi_cell(self._pw, 4.5, val)
        else:
            self.cell(0, 5, val)
            self.ln()

    def _embed_image(self, img_bytes, x, y, max_w, max_h) -> tuple:
        try:
            img = Image.open(io.BytesIO(img_bytes))
            if img.mode in ("RGBA", "P", "LA"):
                img = img.convert("RGB")
            if img.width > 1800 or img.height > 1800:
                img.thumbnail((1800, 1800), Image.LANCZOS)
            ratio = min(max_w / img.width, max_h / img.height)
            w, h = img.width * ratio, img.height * ratio
            buf = io.BytesIO()
            img.save(buf, "JPEG", quality=85)
            buf.seek(0)
            self.image(buf, x=x, y=y, w=w, h=h)
            return w, h
        except Exception:
            return 0, 0

    def add_item(self, item: dict, client: HomeboxClient, maintenance: list = None):
        self.add_page()
        self.item_count += 1
        name = self._t(item.get("name", "Unknown"))
        asset_id = HomeboxClient.format_asset_id(item.get("assetId", 0))

        # Header bar
        bar_y = self.get_y()
        self.set_fill_color(*NAVY)
        self.rect(20, bar_y, self._pw, 14, "F")
        self.set_xy(25, bar_y + 1)
        self.set_font("Helvetica", "B", 12)
        self.set_text_color(*WHITE)
        self.cell(self._pw - 50, 12, name[:48])
        self.set_font("Helvetica", "", 9)
        self.cell(42, 12, asset_id, align="R")
        self.set_y(bar_y + 18)

        # Photo + Details
        attachments = item.get("attachments") or []
        photos = [a for a in attachments if a.get("type") == "photo"]
        primary = next((a for a in photos if a.get("primary")), photos[0] if photos else None)
        section_y = self.get_y()
        photo_w, photo_h = 0, 0
        if primary and primary.get("id"):
            try:
                img_data = client.get_attachment(item["id"], primary["id"])
                photo_w, photo_h = self._embed_image(img_data, 20, section_y, 65, 55)
            except Exception:
                pass

        dx = 20 + (70 if photo_w > 0 else 0)
        self.set_xy(dx, section_y)
        lw, col_w = 32, self._pw - (70 if photo_w > 0 else 0)
        fields = [
            ("Location",     self._loc_name(item)),
            ("Quantity",     str(item.get("quantity", 1)) if item.get("quantity", 1) != 1 else ""),
            ("Serial #",    item.get("serialNumber")),
            ("Model #",     item.get("modelNumber")),
            ("Manufacturer", item.get("manufacturer")),
            ("Insured",     "Yes" if item.get("insured") else "No"),
        ]
        tags = item.get("tags") or []
        if tags:
            fields.append(("Tags", ", ".join(self._t(t.get("name", "")) for t in tags)))
        for label, value in fields:
            if not value:
                continue
            self.set_x(dx)
            self.set_font("Helvetica", "B", 8)
            self.set_text_color(*GRAY_500)
            self.cell(lw, 5, self._t(label))
            self.set_font("Helvetica", "", 8)
            self.set_text_color(*GRAY_700)
            self.cell(col_w - lw, 5, self._t(str(value))[:40])
            self.ln()
        self.set_y(max(self.get_y(), section_y + photo_h) + 5)
        self.set_x(20)

        # Description
        desc = item.get("description")
        if desc:
            self._heading("Description")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(*GRAY_700)
            self.multi_cell(self._pw, 4.5, self._t(desc))
            self.ln(1)

        # Purchase
        pp = float(item.get("purchasePrice") or 0)
        self.total_value += pp
        purchase = [
            ("Purchased From", item.get("purchaseFrom")),
            ("Purchase Date",  self._date(item.get("purchaseTime"))),
            ("Purchase Price", self._price(pp) if pp else ""),
        ]
        if any(v for _, v in purchase):
            self._heading("Purchase Information")
            for l, v in purchase:
                self._row(l, v)

        # Warranty
        warranty = [
            ("Lifetime Warranty", "Yes" if item.get("lifetimeWarranty") else ""),
            ("Expires",           self._date(item.get("warrantyExpires"))),
            ("Details",           item.get("warrantyDetails")),
        ]
        if any(v for _, v in warranty):
            self._heading("Warranty")
            for l, v in warranty:
                self._row(l, v)

        # Custom fields
        custom = item.get("fields") or []
        if custom:
            self._heading("Additional Details")
            for f in custom:
                fn = f.get("name", "")
                ft = f.get("type", "text")
                if ft == "boolean":
                    fv = "Yes" if f.get("booleanValue") else "No"
                elif ft == "number":
                    fv = str(f.get("numberValue", ""))
                else:
                    fv = f.get("textValue", "")
                self._row(fn, fv)

        # Notes
        notes = item.get("notes")
        if notes:
            self._heading("Notes")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(*GRAY_700)
            self.multi_cell(self._pw, 4.5, self._t(notes))
            self.ln(1)

        # Additional photos
        other_photos = [p for p in photos if p != primary]
        if other_photos:
            self._heading("Additional Photos")
            x_pos, row_h = 20, 0
            for photo in other_photos[:8]:
                if not photo.get("id"):
                    continue
                try:
                    img_data = client.get_attachment(item["id"], photo["id"])
                    if x_pos + 48 > self.w - 20:
                        self.set_y(self.get_y() + row_h + 3)
                        x_pos, row_h = 20, 0
                    if self.get_y() + 42 > self.h - 25:
                        self.add_page()
                        x_pos, row_h = 20, 0
                    pw, ph = self._embed_image(img_data, x_pos, self.get_y(), 45, 38)
                    if pw > 0:
                        x_pos += pw + 5
                        row_h = max(row_h, ph)
                except Exception:
                    continue
            if row_h > 0:
                self.set_y(self.get_y() + row_h + 4)

        # Receipts
        receipts = [a for a in attachments if a.get("type") == "receipt"]
        if receipts:
            self._heading("Receipts")
            for rcpt in receipts:
                if not rcpt.get("id"):
                    continue
                try:
                    img_data = client.get_attachment(item["id"], rcpt["id"])
                    if self.get_y() + 50 > self.h - 25:
                        self.add_page()
                    rw, rh = self._embed_image(img_data, 20, self.get_y(), self._pw, 90)
                    if rh > 0:
                        self.set_y(self.get_y() + rh + 2)
                        self.set_font("Helvetica", "I", 7)
                        self.set_text_color(*GRAY_500)
                        self.cell(0, 4, self._t(rcpt.get("title") or "Receipt"))
                        self.ln(4)
                except Exception:
                    continue

        # Maintenance
        if maintenance:
            entries = maintenance if isinstance(maintenance, list) else []
            if entries:
                self._heading("Maintenance History")
                mc = [38, 52, 28, 28, 24]
                mh = ["Task", "Description", "Scheduled", "Completed", "Cost"]
                self.set_fill_color(*NAVY)
                self.set_text_color(*WHITE)
                self.set_font("Helvetica", "B", 7)
                for i, h in enumerate(mh):
                    self.cell(mc[i], 6, f" {h}", fill=True)
                self.ln()
                self.set_font("Helvetica", "", 7)
                for idx, entry in enumerate(entries):
                    if self.get_y() > self.h - 25:
                        self.add_page()
                    self.set_fill_color(*(GRAY_50 if idx % 2 == 0 else WHITE))
                    self.set_text_color(*GRAY_700)
                    self.cell(mc[0], 5, f" {self._t(entry.get('name', ''))[:20]}", fill=True)
                    self.cell(mc[1], 5, f" {self._t(entry.get('description', ''))[:28]}", fill=True)
                    self.cell(mc[2], 5, f" {self._date(entry.get('scheduledDate'))}", fill=True)
                    self.cell(mc[3], 5, f" {self._date(entry.get('completedDate'))}", fill=True)
                    cost = float(entry.get("cost") or 0)
                    self.cell(mc[4], 5, self._price(cost) if cost else " --", fill=True, align="R")
                    self.ln()

        # Sold info
        sold = [
            ("Sold To",    item.get("soldTo")),
            ("Sold Date",  self._date(item.get("soldTime"))),
            ("Sold Price", self._price(float(item.get("soldPrice") or 0)) if float(item.get("soldPrice") or 0) else ""),
            ("Notes",      item.get("soldNotes")),
        ]
        if any(v for _, v in sold):
            self._heading("Sold Information")
            for l, v in sold:
                self._row(l, v)


# ═══════════════════════════════════════════════════════════════════════════════
# Configuration
# ═══════════════════════════════════════════════════════════════════════════════
def load_config() -> dict:
    try:
        return json.loads(_get_config_path().read_text())
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_config(cfg: dict):
    _get_config_path().write_text(json.dumps(cfg, indent=2))


# ═══════════════════════════════════════════════════════════════════════════════
# GUI Application
# ═══════════════════════════════════════════════════════════════════════════════
class HomeboxExportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Homebox Inventory Export")
        self.geometry("860x620")
        self.minsize(750, 520)

        self.client = None
        self.all_items = []           # items currently shown in tree
        self._progress_queue = queue.Queue()
        self._exporting = False

        # Apply theme
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")

        # Custom styles
        style.configure("Navy.TButton", font=("Segoe UI", 9, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self._build_ui()
        self._load_config()

    # ── Build UI ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Main container with padding
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # ── Connection frame ──
        conn = ttk.LabelFrame(main, text="  Connection  ", padding=8)
        conn.pack(fill="x", pady=(0, 6))

        row1 = ttk.Frame(conn)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="Homebox URL:").pack(side="left")
        self.url_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.url_var, width=32).pack(side="left", padx=(5, 15))
        ttk.Label(row1, text="Username:").pack(side="left")
        self.user_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.user_var, width=22).pack(side="left", padx=(5, 15))

        row2 = ttk.Frame(conn)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="Password:").pack(side="left")
        self.pass_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.pass_var, width=20, show="*").pack(side="left", padx=(14, 15))
        ttk.Label(row2, text="Your Name:").pack(side="left")
        self.owner_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.owner_var, width=22).pack(side="left", padx=(5, 15))

        self.connect_btn = ttk.Button(row2, text="Connect", command=self._connect, style="Navy.TButton")
        self.connect_btn.pack(side="right", padx=(10, 0))
        self.conn_status = tk.StringVar(value="Not connected")
        ttk.Label(row2, textvariable=self.conn_status, foreground="gray").pack(side="right")

        # ── Search frame ──
        search = ttk.LabelFrame(main, text="  Find Items  ", padding=8)
        search.pack(fill="x", pady=(0, 6))

        s_row1 = ttk.Frame(search)
        s_row1.pack(fill="x", pady=2)
        ttk.Label(s_row1, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(s_row1, textvariable=self.search_var, width=30)
        self.search_entry.pack(side="left", padx=(5, 5))
        self.search_entry.bind("<Return>", lambda e: self._search())
        ttk.Button(s_row1, text="Search", command=self._search).pack(side="left", padx=2)
        ttk.Button(s_row1, text="Load All", command=self._load_all).pack(side="left", padx=2)

        s_row2 = ttk.Frame(search)
        s_row2.pack(fill="x", pady=2)
        ttk.Label(s_row2, text="Asset IDs:").pack(side="left")
        self.asset_id_var = tk.StringVar()
        self.asset_entry = ttk.Entry(s_row2, textvariable=self.asset_id_var, width=40)
        self.asset_entry.pack(side="left", padx=(5, 5))
        self.asset_entry.bind("<Return>", lambda e: self._find_by_ids())
        ttk.Button(s_row2, text="Find", command=self._find_by_ids).pack(side="left", padx=2)
        ttk.Label(s_row2, text="(comma-separated, e.g. 000-001-234, 000-002-100)",
                  foreground="gray", font=("Segoe UI", 8)).pack(side="left", padx=8)

        # ── Items treeview ──
        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill="both", expand=True, pady=(0, 6))

        columns = ("asset_id", "name", "location", "value", "insured")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                 selectmode="extended")
        self.tree.heading("asset_id", text="Asset ID", command=lambda: self._sort("asset_id"))
        self.tree.heading("name", text="Name", command=lambda: self._sort("name"))
        self.tree.heading("location", text="Location", command=lambda: self._sort("location"))
        self.tree.heading("value", text="Value", command=lambda: self._sort("value"))
        self.tree.heading("insured", text="Insured", command=lambda: self._sort("insured"))

        self.tree.column("asset_id", width=100, minwidth=90)
        self.tree.column("name", width=280, minwidth=150)
        self.tree.column("location", width=160, minwidth=80)
        self.tree.column("value", width=90, minwidth=70, anchor="e")
        self.tree.column("insured", width=70, minwidth=60, anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Tag for alternating row colors
        self.tree.tag_configure("even", background="#f9fafb")
        self.tree.tag_configure("odd", background="#ffffff")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── Action frame ──
        action = ttk.LabelFrame(main, text="  Export  ", padding=8)
        action.pack(fill="x")

        a_row1 = ttk.Frame(action)
        a_row1.pack(fill="x", pady=2)
        self.items_label = tk.StringVar(value="Items: 0")
        self.sel_label = tk.StringVar(value="Selected: 0")
        ttk.Label(a_row1, textvariable=self.items_label, style="Header.TLabel").pack(side="left", padx=(0, 20))
        ttk.Label(a_row1, textvariable=self.sel_label, style="Header.TLabel").pack(side="left", padx=(0, 20))

        ttk.Button(a_row1, text="Select All", command=self._select_all).pack(side="right", padx=2)
        ttk.Button(a_row1, text="Clear Selection", command=self._clear_selection).pack(side="right", padx=2)

        a_row2 = ttk.Frame(action)
        a_row2.pack(fill="x", pady=(4, 2))

        self.export_sel_btn = ttk.Button(a_row2, text="Export Selected to PDF",
                                          command=self._export_selected, style="Navy.TButton")
        self.export_sel_btn.pack(side="left", padx=(0, 8))
        self.export_all_btn = ttk.Button(a_row2, text="Export All to PDF",
                                          command=self._export_all, style="Navy.TButton")
        self.export_all_btn.pack(side="left")

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(a_row2, textvariable=self.status_var, style="Status.TLabel",
                  foreground="gray").pack(side="right")

        a_row3 = ttk.Frame(action)
        a_row3.pack(fill="x", pady=(2, 0))
        self.progress = ttk.Progressbar(a_row3, mode="determinate", length=400)
        self.progress.pack(fill="x")

    # ── Config ────────────────────────────────────────────────────────────────

    def _load_config(self):
        cfg = load_config()
        if cfg.get("url"):
            self.url_var.set(cfg["url"])
        if cfg.get("username"):
            self.user_var.set(cfg["username"])
        if cfg.get("owner"):
            self.owner_var.set(cfg["owner"])

    def _save_config(self):
        save_config({
            "url": self.url_var.get(),
            "username": self.user_var.get(),
            "owner": self.owner_var.get(),
        })

    # ── Connect ───────────────────────────────────────────────────────────────

    def _connect(self):
        url = self.url_var.get().strip()
        user = self.user_var.get().strip()
        pw = self.pass_var.get()

        if not url or not user or not pw:
            messagebox.showwarning("Missing Info", "Please fill in URL, username, and password.")
            return

        self.conn_status.set("Connecting...")
        self.connect_btn.configure(state="disabled")
        self.update_idletasks()

        try:
            self.client = HomeboxClient(url)
            self.client.login(user, pw)
            self.conn_status.set("Connected!")
            self._save_config()
            self.status_var.set("Connected — search or load items")
            # Auto-load all items
            self._load_all()
        except requests.exceptions.ConnectionError:
            self.conn_status.set("Connection failed")
            messagebox.showerror("Connection Error", f"Cannot reach {url}\nCheck the URL and make sure Homebox is running.")
            self.client = None
        except requests.exceptions.HTTPError:
            self.conn_status.set("Login failed")
            messagebox.showerror("Login Failed", "Check your username and password.")
            self.client = None
        finally:
            self.connect_btn.configure(state="normal")

    # ── Search / Load ─────────────────────────────────────────────────────────

    def _require_client(self) -> bool:
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect to Homebox first.")
            return False
        return True

    def _search(self):
        if not self._require_client():
            return
        q = self.search_var.get().strip()
        if not q:
            return
        self.status_var.set(f"Searching for '{q}'...")
        self.update_idletasks()
        try:
            items = self.client.get_items(query=q)
            self.all_items = items
            self._populate_tree(items)
            self.status_var.set(f"Found {len(items)} item(s)")
        except Exception as e:
            messagebox.showerror("Search Error", str(e))
            self.status_var.set("Search failed")

    def _load_all(self):
        if not self._require_client():
            return
        self.status_var.set("Loading all items...")
        self.update_idletasks()
        try:
            items = self.client.get_items()
            self.all_items = items
            self._populate_tree(items)
            self.status_var.set(f"Loaded {len(items)} item(s)")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            self.status_var.set("Load failed")

    def _find_by_ids(self):
        if not self._require_client():
            return
        raw = self.asset_id_var.get().strip()
        if not raw:
            return
        ids = [x.strip() for x in raw.split(",") if x.strip()]
        self.status_var.set("Looking up asset IDs...")
        self.update_idletasks()

        try:
            all_items = self.client.get_items()
            matched = []
            for aid in ids:
                target = aid.replace("-", "").lstrip("0") or "0"
                for it in all_items:
                    item_aid = str(it.get("assetId", "0")).lstrip("0") or "0"
                    if item_aid == target:
                        matched.append(it)
                        break
            self.all_items = matched
            self._populate_tree(matched)
            self.status_var.set(f"Found {len(matched)} of {len(ids)} asset ID(s)")
        except Exception as e:
            messagebox.showerror("Lookup Error", str(e))

    # ── Treeview ──────────────────────────────────────────────────────────────

    def _populate_tree(self, items: list):
        self.tree.delete(*self.tree.get_children())
        for idx, item in enumerate(items):
            aid = HomeboxClient.format_asset_id(item.get("assetId", 0))
            name = item.get("name", "")
            loc_obj = item.get("location")
            loc = loc_obj.get("name", "") if isinstance(loc_obj, dict) else ""
            price = float(item.get("purchasePrice") or 0)
            price_s = f"${price:,.2f}" if price else ""
            ins = "Yes" if item.get("insured") else ""
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", "end", iid=str(idx), values=(aid, name, loc, price_s, ins),
                             tags=(tag,))
        self.items_label.set(f"Items: {len(items)}")
        self.sel_label.set("Selected: 0")

    def _on_select(self, event=None):
        sel = self.tree.selection()
        self.sel_label.set(f"Selected: {len(sel)}")

    def _select_all(self):
        children = self.tree.get_children()
        self.tree.selection_set(children)
        self.sel_label.set(f"Selected: {len(children)}")

    def _clear_selection(self):
        self.tree.selection_remove(self.tree.selection())
        self.sel_label.set("Selected: 0")

    def _sort(self, col):
        """Sort treeview by column header click."""
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children()]
        # Try numeric sort for value column
        if col == "value":
            def sort_key(item):
                v = item[0].replace("$", "").replace(",", "")
                try:
                    return float(v) if v else 0
                except ValueError:
                    return 0
            data.sort(key=sort_key, reverse=True)
        else:
            data.sort(key=lambda t: t[0].lower())
        for idx, (val, k) in enumerate(data):
            self.tree.move(k, "", idx)
            self.tree.item(k, tags=("even" if idx % 2 == 0 else "odd",))

    # ── Export ────────────────────────────────────────────────────────────────

    def _get_selected_items(self) -> list:
        sel = self.tree.selection()
        return [self.all_items[int(iid)] for iid in sel if int(iid) < len(self.all_items)]

    def _export_selected(self):
        if not self._require_client():
            return
        items = self._get_selected_items()
        if not items:
            messagebox.showinfo("No Selection", "Select items in the list first.\n\n"
                                "Tip: Ctrl+Click to select multiple,\n"
                                "Shift+Click to select a range.")
            return
        self._start_export(items)

    def _export_all(self):
        if not self._require_client():
            return
        if not self.all_items:
            messagebox.showinfo("No Items", "Load or search for items first.")
            return
        count = len(self.all_items)
        if not messagebox.askyesno("Export All", f"Export all {count} items to PDF?"):
            return
        self._start_export(self.all_items)

    def _start_export(self, items: list):
        if self._exporting:
            return

        # Ask where to save
        today = datetime.now().strftime("%m-%d-%Y")
        if len(items) == 1:
            aid = HomeboxClient.format_asset_id(items[0].get("assetId", 0))
            default_name = f"HomeBox Asset Export {aid} - {today}.pdf"
        else:
            default_name = f"HomeBox Asset Export - {today}.pdf"
        path = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_name,
        )
        if not path:
            return

        self._exporting = True
        self.export_sel_btn.configure(state="disabled")
        self.export_all_btn.configure(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = len(items)

        thread = threading.Thread(
            target=self._export_worker,
            args=(items, path, self.owner_var.get().strip()),
            daemon=True,
        )
        thread.start()
        self.after(100, self._check_progress)

    def _export_worker(self, items: list, save_path: str, owner: str):
        """Runs in background thread."""
        try:
            pdf = InventoryReport()
            pdf.add_cover(owner)

            # Fetch full details
            full = []
            for i, s in enumerate(items):
                name = s.get("name", "?")
                self._progress_queue.put(("progress", i, name))
                try:
                    detail = self.client.get_item(s["id"])
                    maint = self.client.get_maintenance(s["id"])
                except Exception:
                    detail = s
                    maint = []
                full.append((detail, maint))

            self._progress_queue.put(("progress", len(items), "Building PDF..."))

            # Summary page for multi-item
            if len(full) > 1:
                pdf.add_summary([f[0] for f in full])

            # Item pages
            for detail, maint in full:
                try:
                    pdf.add_item(detail, self.client, maint)
                except Exception:
                    pass

            pdf.output(save_path)
            self._progress_queue.put(("done", save_path, pdf.item_count, pdf.total_value))

        except Exception as e:
            self._progress_queue.put(("error", str(e)))

    def _check_progress(self):
        try:
            while True:
                msg = self._progress_queue.get_nowait()
                if msg[0] == "progress":
                    _, idx, name = msg
                    self.progress["value"] = idx
                    self.status_var.set(f"[{idx + 1}/{int(self.progress['maximum'])}] {name}")
                elif msg[0] == "done":
                    _, path, count, total = msg
                    self.progress["value"] = self.progress["maximum"]
                    self.status_var.set(f"Saved: {Path(path).name}")
                    self._exporting = False
                    self.export_sel_btn.configure(state="normal")
                    self.export_all_btn.configure(state="normal")
                    messagebox.showinfo(
                        "Export Complete",
                        f"PDF saved successfully!\n\n"
                        f"File: {path}\n"
                        f"Items: {count}\n"
                        f"Total Value: ${total:,.2f}"
                    )
                    return
                elif msg[0] == "error":
                    self.status_var.set("Export failed")
                    self._exporting = False
                    self.export_sel_btn.configure(state="normal")
                    self.export_all_btn.configure(state="normal")
                    messagebox.showerror("Export Error", msg[1])
                    return
        except queue.Empty:
            pass

        if self._exporting:
            self.after(100, self._check_progress)


# ═══════════════════════════════════════════════════════════════════════════════
# Entry Point
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = HomeboxExportApp()
    app.mainloop()
