#!/usr/bin/env python3
"""
Homebox Export Tool
Generates professional PDF inventory reports from your Homebox instance.
Designed for insurance documentation, warranty claims, and asset records.
"""

import os
import sys
import io
import json
import getpass
from datetime import datetime
from pathlib import Path

try:
    import requests
except ImportError:
    print("Missing: requests\n  pip install requests")
    sys.exit(1)

try:
    from fpdf import FPDF
except ImportError:
    print("Missing: fpdf2\n  pip install fpdf2")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("Missing: Pillow\n  pip install Pillow")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════════════════════
# Color Palette
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
GREEN_700  = (21, 128, 61)
RED_600    = (220, 38, 38)

CONFIG_PATH = Path.home() / ".homebox_export.json"


# ═══════════════════════════════════════════════════════════════════════════════
# Homebox API Client
# ═══════════════════════════════════════════════════════════════════════════════
class HomeboxClient:
    """Communicates with the Homebox REST API."""

    def __init__(self, base_url: str):
        self.base_url = base_url.rstrip("/")
        self.api = f"{self.base_url}/api/v1"
        self.session = requests.Session()
        self.session.headers["Accept"] = "application/json"

    def login(self, username: str, password: str) -> dict:
        r = self.session.post(f"{self.api}/users/login", json={
            "username": username,
            "password": password,
        })
        r.raise_for_status()
        data = r.json()

        token = data.get("token", "")
        # Homebox may return the token with or without "Bearer " prefix
        if token.startswith("Bearer "):
            self.session.headers["Authorization"] = token
        else:
            self.session.headers["Authorization"] = f"Bearer {token}"

        # Store attachment token cookie if provided
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

    @staticmethod
    def parse_asset_id(formatted: str) -> str:
        """Convert '000-001-234' to '1234' (the integer string form)."""
        try:
            return str(int(formatted.replace("-", "")))
        except (ValueError, TypeError):
            return formatted


# ═══════════════════════════════════════════════════════════════════════════════
# PDF Report Generator
# ═══════════════════════════════════════════════════════════════════════════════
class InventoryReport(FPDF):
    """Generates a professional insurance-grade PDF inventory report."""

    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="Letter")
        self.set_auto_page_break(True, margin=25)
        self.set_margins(20, 20, 20)
        self._pw = self.w - 40          # usable page width (170mm)
        self.owner = ""
        self.total_value = 0.0
        self.item_count = 0
        self._is_cover = True

    # ── Text helpers ──────────────────────────────────────────────────────────

    @staticmethod
    def _t(text) -> str:
        """Sanitize text for Latin-1 PDF fonts."""
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
        if s.startswith("0001") or s == "":
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
        if isinstance(loc, dict):
            return loc.get("name", "")
        return ""

    # ── Page header / footer ──────────────────────────────────────────────────

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
        # Page number (subtract 1 for cover page)
        self.cell(0, 8, f"Page {self.page_no() - 1}", align="C")

    # ── Cover page ────────────────────────────────────────────────────────────

    def add_cover(self, owner: str = ""):
        self.owner = owner
        self._is_cover = True
        self.add_page()

        # Navy banner
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

        # White divider line
        self.set_draw_color(*WHITE)
        self.set_line_width(0.4)
        cx = self.w / 2
        self.line(cx - 25, self.get_y(), cx + 25, self.get_y())
        self.ln(7)

        self.set_font("Helvetica", "I", 11)
        self.cell(0, 7, "For Insurance & Documentation Purposes", align="C")

        # Info below banner
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

    # ── Summary page ──────────────────────────────────────────────────────────

    def add_summary(self, items: list):
        self.add_page()

        self.set_font("Helvetica", "B", 18)
        self.set_text_color(*NAVY)
        self.cell(0, 10, "Inventory Summary")
        self.ln(14)

        # Stat boxes
        total_val = sum(float(it.get("purchasePrice") or 0) for it in items)
        insured = sum(1 for it in items if it.get("insured"))
        stats = [
            ("Total Items", str(len(items))),
            ("Estimated Value", f"${total_val:,.2f}"),
            ("Insured Items", str(insured)),
        ]

        box_w = self._pw / 3 - 3
        start_x = 20
        y = self.get_y()

        for i, (label, value) in enumerate(stats):
            x = start_x + i * (box_w + 4.5)
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

        # Item table
        cols = [26, 60, 38, 24, 22]
        headers = ["Asset ID", "Name", "Location", "Value", "Insured"]

        self._draw_table_header(cols, headers)

        self.set_font("Helvetica", "", 8)
        for idx, item in enumerate(items):
            if self.get_y() > self.h - 30:
                self.add_page()
                self._draw_table_header(cols, headers)
                self.set_font("Helvetica", "", 8)

            bg = GRAY_50 if idx % 2 == 0 else WHITE
            self.set_fill_color(*bg)
            self.set_text_color(*GRAY_700)

            aid = HomeboxClient.format_asset_id(item.get("assetId", 0))
            name = self._t(item.get("name", ""))[:33]
            loc = self._t(self._loc_name(item))[:20]
            price = self._price(item.get("purchasePrice"))
            ins = "Yes" if item.get("insured") else ""

            self.cell(cols[0], 6, f" {aid}", fill=True)
            self.cell(cols[1], 6, f" {name}", fill=True)
            self.cell(cols[2], 6, f" {loc}", fill=True)
            self.cell(cols[3], 6, f"{price} ", fill=True, align="R")
            self.cell(cols[4], 6, ins, fill=True, align="C")
            self.ln()

    def _draw_table_header(self, cols, headers):
        self.set_fill_color(*NAVY)
        self.set_text_color(*WHITE)
        self.set_font("Helvetica", "B", 8)
        for i, h in enumerate(headers):
            al = "R" if h == "Value" else ("C" if h == "Insured" else "L")
            self.cell(cols[i], 7, f" {h}", fill=True, align=al)
        self.ln()

    # ── Section header ────────────────────────────────────────────────────────

    def _heading(self, title):
        if self.get_y() > self.h - 28:
            self.add_page()
        self.ln(2)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*NAVY)
        self.cell(0, 6, self._t(title))
        self.ln()
        y = self.get_y()
        self.set_draw_color(*BLUE_600)
        self.set_line_width(0.3)
        self.line(20, y, self.w - 20, y)
        self.ln(3)

    # ── Detail row ────────────────────────────────────────────────────────────

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

    # ── Embed image safely ────────────────────────────────────────────────────

    def _embed_image(self, img_bytes, x, y, max_w, max_h) -> tuple:
        """Returns (width, height) placed, or (0, 0) on failure."""
        try:
            img = Image.open(io.BytesIO(img_bytes))
            if img.mode in ("RGBA", "P", "LA"):
                img = img.convert("RGB")
            # Downscale very large images
            if img.width > 1800 or img.height > 1800:
                img.thumbnail((1800, 1800), Image.LANCZOS)

            ratio = min(max_w / img.width, max_h / img.height)
            w = img.width * ratio
            h = img.height * ratio

            buf = io.BytesIO()
            img.save(buf, "JPEG", quality=85)
            buf.seek(0)

            self.image(buf, x=x, y=y, w=w, h=h)
            return w, h
        except Exception:
            return 0, 0

    # ── Full item page ────────────────────────────────────────────────────────

    def add_item(self, item: dict, client: HomeboxClient, maintenance: list = None):
        self.add_page()
        self.item_count += 1

        name = self._t(item.get("name", "Unknown"))
        asset_id = HomeboxClient.format_asset_id(item.get("assetId", 0))

        # ── Header bar ──
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

        # ── Photo + Details side-by-side ──
        attachments = item.get("attachments") or []
        photos = [a for a in attachments if a.get("type") == "photo"]
        primary = next((a for a in photos if a.get("primary")), photos[0] if photos else None)

        section_y = self.get_y()
        photo_w, photo_h = 0, 0

        if primary:
            att_id = primary.get("id")
            if att_id:
                try:
                    img_data = client.get_attachment(item["id"], att_id)
                    photo_w, photo_h = self._embed_image(
                        img_data, 20, section_y, 65, 55
                    )
                except Exception:
                    pass

        # Details column (beside photo or full-width)
        dx = 20 + (70 if photo_w > 0 else 0)
        self.set_xy(dx, section_y)
        lw = 32
        col_w = self._pw - (70 if photo_w > 0 else 0)

        fields = [
            ("Location",     self._loc_name(item)),
            ("Quantity",     str(item.get("quantity", 1)) if item.get("quantity", 1) != 1 else ""),
            ("Serial #",    item.get("serialNumber")),
            ("Model #",     item.get("modelNumber")),
            ("Manufacturer", item.get("manufacturer")),
            ("Insured",     "Yes" if item.get("insured") else "No"),
        ]

        # Tags
        tags = item.get("tags") or []
        if tags:
            tag_str = ", ".join(self._t(t.get("name", "")) for t in tags)
            fields.append(("Tags", tag_str))

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

        # Move past the taller of photo or details
        self.set_y(max(self.get_y(), section_y + photo_h) + 5)
        self.set_x(20)

        # ── Description ──
        desc = item.get("description")
        if desc:
            self._heading("Description")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(*GRAY_700)
            self.multi_cell(self._pw, 4.5, self._t(desc))
            self.ln(1)

        # ── Purchase Information ──
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

        # ── Warranty ──
        warranty = [
            ("Lifetime Warranty", "Yes" if item.get("lifetimeWarranty") else ""),
            ("Expires",           self._date(item.get("warrantyExpires"))),
            ("Details",           item.get("warrantyDetails")),
        ]
        if any(v for _, v in warranty):
            self._heading("Warranty")
            for l, v in warranty:
                self._row(l, v)

        # ── Custom fields ──
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

        # ── Notes ──
        notes = item.get("notes")
        if notes:
            self._heading("Notes")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(*GRAY_700)
            self.multi_cell(self._pw, 4.5, self._t(notes))
            self.ln(1)

        # ── Additional Photos ──
        other_photos = [p for p in photos if p != primary]
        if other_photos:
            self._heading("Additional Photos")
            x_pos = 20
            row_h = 0
            for photo in other_photos[:8]:
                att_id = photo.get("id")
                if not att_id:
                    continue
                try:
                    img_data = client.get_attachment(item["id"], att_id)
                    if x_pos + 48 > self.w - 20:
                        self.set_y(self.get_y() + row_h + 3)
                        x_pos = 20
                        row_h = 0
                    if self.get_y() + 42 > self.h - 25:
                        self.add_page()
                        x_pos = 20
                        row_h = 0
                    pw, ph = self._embed_image(img_data, x_pos, self.get_y(), 45, 38)
                    if pw > 0:
                        x_pos += pw + 5
                        row_h = max(row_h, ph)
                except Exception:
                    continue
            if row_h > 0:
                self.set_y(self.get_y() + row_h + 4)

        # ── Receipt images ──
        receipts = [a for a in attachments if a.get("type") == "receipt"]
        if receipts:
            self._heading("Receipts")
            for rcpt in receipts:
                att_id = rcpt.get("id")
                if not att_id:
                    continue
                try:
                    img_data = client.get_attachment(item["id"], att_id)
                    if self.get_y() + 50 > self.h - 25:
                        self.add_page()
                    rw, rh = self._embed_image(
                        img_data, 20, self.get_y(), self._pw, 90
                    )
                    if rh > 0:
                        self.set_y(self.get_y() + rh + 2)
                        title = rcpt.get("title") or "Receipt"
                        self.set_font("Helvetica", "I", 7)
                        self.set_text_color(*GRAY_500)
                        self.cell(0, 4, self._t(title))
                        self.ln(4)
                except Exception:
                    continue

        # ── Maintenance History ──
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
                    bg = GRAY_50 if idx % 2 == 0 else WHITE
                    self.set_fill_color(*bg)
                    self.set_text_color(*GRAY_700)

                    self.cell(mc[0], 5, f" {self._t(entry.get('name', ''))[:20]}", fill=True)
                    self.cell(mc[1], 5, f" {self._t(entry.get('description', ''))[:28]}", fill=True)
                    self.cell(mc[2], 5, f" {self._date(entry.get('scheduledDate'))}", fill=True)
                    self.cell(mc[3], 5, f" {self._date(entry.get('completedDate'))}", fill=True)
                    cost = float(entry.get("cost") or 0)
                    self.cell(mc[4], 5, self._price(cost) if cost else " --", fill=True, align="R")
                    self.ln()

        # ── Sold Information (if applicable) ──
        sold_time = self._date(item.get("soldTime"))
        sold_to = item.get("soldTo")
        sold_price = float(item.get("soldPrice") or 0)
        sold_notes = item.get("soldNotes")
        sold = [
            ("Sold To",    sold_to),
            ("Sold Date",  sold_time),
            ("Sold Price", self._price(sold_price) if sold_price else ""),
            ("Notes",      sold_notes),
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
        return json.loads(CONFIG_PATH.read_text())
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_config(cfg: dict):
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2))


# ═══════════════════════════════════════════════════════════════════════════════
# CLI Display Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def clear():
    os.system("cls" if os.name == "nt" else "clear")


def banner():
    print()
    print("  " + "=" * 48)
    print("      HOMEBOX  INVENTORY  EXPORT  TOOL")
    print("      Professional PDF Report Generator")
    print("  " + "=" * 48)
    print()


def display_items(items: list):
    if not items:
        print("  No items found.\n")
        return
    print()
    print(f"  {'Asset ID':<14} {'Name':<32} {'Location':<18} {'Value':>10}")
    print(f"  {'─' * 13}  {'─' * 31}  {'─' * 17}  {'─' * 10}")
    for item in items:
        aid = HomeboxClient.format_asset_id(item.get("assetId", 0))
        name = item.get("name", "?")[:30]
        loc_obj = item.get("location")
        loc = loc_obj.get("name", "")[:16] if isinstance(loc_obj, dict) else ""
        price = float(item.get("purchasePrice") or 0)
        ps = f"${price:,.2f}" if price else ""
        print(f"  {aid:<14} {name:<32} {loc:<18} {ps:>10}")
    print()


# ═══════════════════════════════════════════════════════════════════════════════
# PDF Generation
# ═══════════════════════════════════════════════════════════════════════════════

def generate_pdf(client: HomeboxClient, items_summary: list, owner: str) -> Path:
    count = len(items_summary)
    print(f"\n  Generating report for {count} item(s)...\n")

    pdf = InventoryReport()
    pdf.add_cover(owner)

    # Fetch full details for each item
    full = []
    for i, s in enumerate(items_summary):
        iid = s.get("id")
        name = s.get("name", "?")
        pct = int((i + 1) / count * 100)
        print(f"  [{i+1}/{count}] {pct:>3}%  Fetching: {name}")
        try:
            detail = client.get_item(iid)
            maint = client.get_maintenance(iid)
        except Exception as e:
            print(f"         Warning: {e}")
            detail = s
            maint = []
        full.append((detail, maint))

    # Summary page for multi-item exports
    if len(full) > 1:
        pdf.add_summary([f[0] for f in full])

    # Individual item pages
    for detail, maint in full:
        try:
            pdf.add_item(detail, client, maint)
        except Exception as e:
            print(f"  Warning: page error for {detail.get('name', '?')}: {e}")

    # Save
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    filename = f"homebox_inventory_{ts}.pdf"
    out_path = Path.cwd() / filename
    pdf.output(str(out_path))

    print(f"\n  {'=' * 44}")
    print(f"  PDF saved:    {out_path}")
    print(f"  Items:        {pdf.item_count}")
    print(f"  Total value:  ${pdf.total_value:,.2f}")
    print(f"  {'=' * 44}")
    return out_path


# ═══════════════════════════════════════════════════════════════════════════════
# Asset ID Matching
# ═══════════════════════════════════════════════════════════════════════════════

def find_items_by_asset_ids(client: HomeboxClient, asset_ids: list) -> list:
    """Look up items by formatted asset IDs (e.g., '000-001-234')."""
    all_items = client.get_items()
    matched = []

    for aid in asset_ids:
        aid = aid.strip()
        if not aid:
            continue
        target = aid.replace("-", "").lstrip("0") or "0"
        found = None
        for it in all_items:
            item_aid = str(it.get("assetId", "0")).lstrip("0") or "0"
            if item_aid == target:
                found = it
                break
        if found:
            print(f"  Found:     [{aid}] {found.get('name')}")
            matched.append(found)
        else:
            print(f"  Not found: [{aid}]")

    return matched


# ═══════════════════════════════════════════════════════════════════════════════
# Main Interactive CLI
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    clear()
    banner()

    cfg = load_config()

    # ── Connection setup ──
    default_url = cfg.get("url", "http://192.168.0.100:3100")
    url = input(f"  Homebox URL [{default_url}]: ").strip() or default_url

    default_user = cfg.get("username", "")
    user_prompt = f"  Username [{default_user}]: " if default_user else "  Username: "
    username = input(user_prompt).strip() or default_user

    password = getpass.getpass("  Password: ")

    print(f"\n  Connecting to {url} ...")
    client = HomeboxClient(url)
    try:
        client.login(username, password)
    except requests.exceptions.ConnectionError:
        print(f"\n  ERROR: Cannot reach {url}")
        print("  Check the URL and make sure Homebox is running.\n")
        sys.exit(1)
    except requests.exceptions.HTTPError as e:
        status = e.response.status_code if e.response is not None else "?"
        print(f"\n  ERROR: Login failed (HTTP {status})")
        print("  Check your username and password.\n")
        sys.exit(1)

    print("  Connected!\n")

    # ── Owner name ──
    default_owner = cfg.get("owner", "")
    owner_prompt = f"  Your name (for cover page) [{default_owner}]: " if default_owner else "  Your name (for cover page): "
    owner = input(owner_prompt).strip() or default_owner

    # Save settings (never the password)
    save_config({"url": url, "username": username, "owner": owner})

    # ── Main menu loop ──
    while True:
        print()
        print("  " + "-" * 42)
        print("  1)  Search items")
        print("  2)  Export by Asset ID(s)")
        print("  3)  Export ALL items")
        print("  4)  Quit")
        print("  " + "-" * 42)

        choice = input("  > ").strip()

        if choice == "1":
            q = input("\n  Search: ").strip()
            if not q:
                continue
            print("  Searching...")
            results = client.get_items(query=q)
            display_items(results)
            if results:
                ans = input(f"  Export these {len(results)} item(s) to PDF? (y/n): ").strip().lower()
                if ans == "y":
                    generate_pdf(client, results, owner)

        elif choice == "2":
            print("\n  Enter Asset IDs separated by commas")
            print("  Example: 000-001-234, 000-001-235")
            raw = input("  Asset IDs: ").strip()
            if not raw:
                continue
            ids = [x.strip() for x in raw.split(",") if x.strip()]
            print()
            matched = find_items_by_asset_ids(client, ids)
            if matched:
                display_items(matched)
                ans = input(f"  Export {len(matched)} item(s) to PDF? (y/n): ").strip().lower()
                if ans == "y":
                    generate_pdf(client, matched, owner)
            else:
                print("  No matching items found.\n")

        elif choice == "3":
            print("\n  Fetching all items...")
            items = client.get_items()
            print(f"  Found {len(items)} items.")
            display_items(items[:15])
            if len(items) > 15:
                print(f"  ... and {len(items) - 15} more\n")
            if items:
                ans = input(f"  Export ALL {len(items)} items to PDF? (y/n): ").strip().lower()
                if ans == "y":
                    generate_pdf(client, items, owner)

        elif choice == "4":
            print("\n  Goodbye!\n")
            break

        else:
            print("  Invalid choice.\n")


if __name__ == "__main__":
    main()
