#!/usr/bin/env python3
"""
Homebox Export Tool — GUI Application
Generates professional PDF inventory reports from your Homebox instance.
Authenticates with a Homebox API key stored in Windows Credential Manager.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import queue
from datetime import datetime
from pathlib import Path

from hbx import config
from hbx.client import (
    ASSET_ID_RE, HomeboxClient, HomeboxError,
    format_asset_id, in_asset_range, location_ids,
)
from hbx.report import InventoryReport, loc_name, fmt_price


class HomeboxExportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Homebox Inventory Export")
        self.geometry("900x680")
        self.minsize(780, 560)

        self.client = None
        self.all_items = []           # items currently shown in tree
        self._progress_queue = queue.Queue()
        self._exporting = False
        self._location_choices = [("Any location", None)]
        self._tag_vars = []           # (BooleanVar, tag_id, tag_name)

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")

        style.configure("Navy.TButton", font=("Segoe UI", 9, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self._build_ui()
        self._load_config()

    # ── Build UI ──────────────────────────────────────────────────────────────

    def _build_ui(self):
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
        ttk.Label(row1, text="Your Name:").pack(side="left")
        self.owner_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.owner_var, width=22).pack(side="left", padx=(5, 15))

        row2 = ttk.Frame(conn)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="API Key:").pack(side="left")
        self.key_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.key_var, width=40, show="*").pack(side="left", padx=(5, 5))
        self.key_status = tk.StringVar(value="")
        ttk.Label(row2, textvariable=self.key_status, foreground="gray").pack(side="left", padx=(0, 10))

        self.connect_btn = ttk.Button(row2, text="Connect", command=self._connect, style="Navy.TButton")
        self.connect_btn.pack(side="right", padx=(6, 0))
        ttk.Button(row2, text="Test Connection", command=self._test_connection).pack(side="right", padx=(10, 0))
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

        ttk.Label(s_row1, text="Asset IDs:").pack(side="left", padx=(20, 0))
        self.asset_id_var = tk.StringVar()
        self.asset_entry = ttk.Entry(s_row1, textvariable=self.asset_id_var, width=26)
        self.asset_entry.pack(side="left", padx=(5, 5))
        self.asset_entry.bind("<Return>", lambda e: self._find_by_ids())
        ttk.Button(s_row1, text="Find", command=self._find_by_ids).pack(side="left", padx=2)

        # ── Filters row ──
        s_row2 = ttk.Frame(search)
        s_row2.pack(fill="x", pady=(6, 2))
        ttk.Label(s_row2, text="Location:").pack(side="left")
        self.location_combo = ttk.Combobox(s_row2, state="readonly", width=28,
                                           values=["Any location"])
        self.location_combo.current(0)
        self.location_combo.pack(side="left", padx=(5, 15))

        self.tags_btn = ttk.Menubutton(s_row2, text="Tags")
        self.tags_menu = tk.Menu(self.tags_btn, tearoff=False)
        self.tags_btn.configure(menu=self.tags_menu)
        self.tags_btn.pack(side="left", padx=(0, 15))

        ttk.Label(s_row2, text="Asset range:").pack(side="left")
        self.range_from_var = tk.StringVar()
        self.range_to_var = tk.StringVar()
        ttk.Entry(s_row2, textvariable=self.range_from_var, width=9).pack(side="left", padx=(5, 2))
        ttk.Label(s_row2, text="to").pack(side="left")
        ttk.Entry(s_row2, textvariable=self.range_to_var, width=9).pack(side="left", padx=(2, 10))
        ttk.Label(s_row2, text="(###-###)", foreground="gray",
                  font=("Segoe UI", 8)).pack(side="left")
        ttk.Button(s_row2, text="Clear Filters", command=self._clear_filters).pack(side="right")

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
        cfg = config.load_config()
        if cfg.get("homebox_url"):
            self.url_var.set(cfg["homebox_url"])
        if cfg.get("owner"):
            self.owner_var.set(cfg["owner"])
        if config.load_api_key():
            self.key_status.set("(saved key found)")

    def _save_config(self):
        config.save_config({
            "homebox_url": self.url_var.get().strip(),
            "owner": self.owner_var.get().strip(),
        })

    # ── Connect ───────────────────────────────────────────────────────────────

    def _current_key(self) -> str:
        return self.key_var.get().strip() or config.load_api_key()

    def _test_connection(self):
        url = self.url_var.get().strip()
        key = self._current_key()
        if not url or not key:
            messagebox.showwarning("Missing Info", "Enter the Homebox URL and API key first.")
            return
        client = HomeboxClient(url, key, max_retries=1)
        ok, msg = client.test_connection()
        (messagebox.showinfo if ok else messagebox.showerror)("Test Connection", msg)

    def _connect(self):
        url = self.url_var.get().strip()
        key = self._current_key()

        if not url or not key:
            messagebox.showwarning("Missing Info", "Please fill in the URL and API key.")
            return

        self.conn_status.set("Connecting...")
        self.connect_btn.configure(state="disabled")
        self.update_idletasks()

        try:
            probe = HomeboxClient(url, key, max_retries=1)
            ok, msg = probe.test_connection()
            if not ok:
                self.conn_status.set("Connection failed")
                messagebox.showerror("Connection Failed", msg)
                self.client = None
                return
            self.client = HomeboxClient(url, key)
            self.conn_status.set("Connected!")
            self._save_config()
            if self.key_var.get().strip():
                try:
                    config.save_api_key(self.key_var.get().strip())
                    self.key_var.set("")
                    self.key_status.set("(saved key found)")
                except Exception:
                    messagebox.showwarning(
                        "Credential Manager",
                        "Connected, but the API key could not be saved to "
                        "Windows Credential Manager. You will need to enter it again next time.")
            self.status_var.set("Connected — search or load items")
            self._load_filter_data()
            self._load_all()
        finally:
            self.connect_btn.configure(state="normal")

    # ── Filters ───────────────────────────────────────────────────────────────

    def _load_filter_data(self):
        try:
            tags = self.client.list_tags()
            tree = self.client.location_tree()
        except HomeboxError as e:
            self.status_var.set(f"Filter load failed: {e}")
            return

        self._location_choices = [("Any location", None)]

        def walk(nodes, depth):
            for node in nodes:
                if node.get("type") != "location":
                    continue
                label = "    " * depth + node.get("name", "?")
                self._location_choices.append((label, location_ids(node)))
                walk(node.get("children", []), depth + 1)

        walk(tree, 0)
        self.location_combo.configure(values=[c[0] for c in self._location_choices])
        self.location_combo.current(0)

        self.tags_menu.delete(0, "end")
        self._tag_vars = []
        for tag in tags:
            var = tk.BooleanVar(value=False)
            self._tag_vars.append((var, tag.get("id"), tag.get("name", "?")))
            self.tags_menu.add_checkbutton(
                label=tag.get("name", "?"), variable=var,
                command=self._on_tags_changed)

    def _selected_tag_ids(self):
        return [tid for var, tid, _name in self._tag_vars if var.get()]

    def _on_tags_changed(self):
        count = len(self._selected_tag_ids())
        self.tags_btn.configure(text=f"Tags ({count})" if count else "Tags")

    def _selected_parent_ids(self):
        idx = self.location_combo.current()
        if 0 <= idx < len(self._location_choices):
            return self._location_choices[idx][1]
        return None

    def _valid_range(self):
        lo = self.range_from_var.get().strip()
        hi = self.range_to_var.get().strip()
        lo = lo if ASSET_ID_RE.match(lo) else ""
        hi = hi if ASSET_ID_RE.match(hi) else ""
        return lo, hi

    def _clear_filters(self):
        self.location_combo.current(0)
        for var, _tid, _name in self._tag_vars:
            var.set(False)
        self.tags_btn.configure(text="Tags")
        self.range_from_var.set("")
        self.range_to_var.set("")

    # ── Search / Load ─────────────────────────────────────────────────────────

    def _require_client(self) -> bool:
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect to Homebox first.")
            return False
        return True

    def _run_search(self, query):
        tags = self._selected_tag_ids() or None
        parent_ids = self._selected_parent_ids()
        lo, hi = self._valid_range()
        items, truncated = self.client.search_items_all(
            query, tags=tags, parent_ids=parent_ids)
        if lo or hi:
            items = [it for it in items
                     if in_asset_range(str(it.get("assetId", "")), lo, hi)]
        return items, truncated

    def _search(self):
        if not self._require_client():
            return
        q = self.search_var.get().strip()
        self.status_var.set(f"Searching for '{q}'..." if q else "Searching...")
        self.update_idletasks()
        try:
            items, truncated = self._run_search(q)
            self.all_items = items
            self._populate_tree(items)
            note = " (truncated)" if truncated else ""
            self.status_var.set(f"Found {len(items)} item(s){note}")
        except HomeboxError as e:
            messagebox.showerror("Search Error", str(e))
            self.status_var.set("Search failed")

    def _load_all(self):
        if not self._require_client():
            return
        self.status_var.set("Loading all items...")
        self.update_idletasks()
        try:
            items, truncated = self._run_search("")
            self.all_items = items
            self._populate_tree(items)
            note = " (truncated)" if truncated else ""
            self.status_var.set(f"Loaded {len(items)} item(s){note}")
        except HomeboxError as e:
            messagebox.showerror("Load Error", str(e))
            self.status_var.set("Load failed")

    def _find_by_ids(self):
        if not self._require_client():
            return
        raw = self.asset_id_var.get().strip()
        if not raw:
            return
        ids = [format_asset_id(x) for x in raw.split(",") if x.strip()]
        self.status_var.set("Looking up asset IDs...")
        self.update_idletasks()

        try:
            matched = []
            missing = []
            for aid in ids:
                item = self.client.get_asset(aid)
                (matched if item else missing).append(item or aid)
            self.all_items = matched
            self._populate_tree(matched)
            status = f"Found {len(matched)} of {len(ids)} asset ID(s)"
            if missing:
                status += f" (not found: {', '.join(missing)})"
            self.status_var.set(status)
        except HomeboxError as e:
            messagebox.showerror("Lookup Error", str(e))

    # ── Treeview ──────────────────────────────────────────────────────────────

    def _populate_tree(self, items: list):
        self.tree.delete(*self.tree.get_children())
        for idx, item in enumerate(items):
            aid = format_asset_id(item.get("assetId", ""))
            name = item.get("name", "")
            loc = loc_name(item)
            price_s = fmt_price(item.get("purchasePrice"))
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

        today = datetime.now().strftime("%m-%d-%Y")
        if len(items) == 1:
            aid = format_asset_id(items[0].get("assetId", ""))
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

            full = []
            for i, s in enumerate(items):
                name = s.get("name", "?")
                self._progress_queue.put(("progress", i, name))
                try:
                    detail = self.client.get_entity(s["id"])
                    maint = self.client.get_maintenance(s["id"])
                except HomeboxError:
                    detail = s
                    maint = []
                full.append((detail, maint))

            self._progress_queue.put(("progress", len(items), "Building PDF..."))

            if len(full) > 1:
                pdf.add_summary([f[0] for f in full])

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
