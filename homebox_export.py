#!/usr/bin/env python3
"""
Homebox Export Tool — CLI
Generates professional PDF inventory reports from your Homebox instance.
Authenticates with a Homebox API key stored in Windows Credential Manager.
"""

import getpass
import os
import sys
from datetime import datetime
from pathlib import Path

from hbx import config
from hbx.client import (
    AuthError, HomeboxClient, HomeboxError, format_asset_id,
)
from hbx.report import InventoryReport, fmt_price, loc_name


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
    print(f"  {'Asset ID':<12} {'Name':<32} {'Location':<18} {'Value':>10}")
    print(f"  {'-' * 11}  {'-' * 31}  {'-' * 17}  {'-' * 10}")
    for item in items:
        aid = format_asset_id(item.get("assetId", ""))
        name = item.get("name", "?")[:30]
        loc = loc_name(item)[:16]
        ps = fmt_price(item.get("purchasePrice"))
        print(f"  {aid:<12} {name:<32} {loc:<18} {ps:>10}")
    print()


def generate_pdf(client: HomeboxClient, items_summary: list, owner: str) -> Path:
    count = len(items_summary)
    print(f"\n  Generating report for {count} item(s)...\n")

    pdf = InventoryReport()
    pdf.add_cover(owner)

    full = []
    for i, s in enumerate(items_summary):
        iid = s.get("id")
        name = s.get("name", "?")
        pct = int((i + 1) / count * 100)
        print(f"  [{i+1}/{count}] {pct:>3}%  Fetching: {name}")
        try:
            detail = client.get_entity(iid)
            maint = client.get_maintenance(iid)
        except HomeboxError as e:
            print(f"         Warning: {e}")
            detail = s
            maint = []
        full.append((detail, maint))

    if len(full) > 1:
        pdf.add_summary([f[0] for f in full])

    for detail, maint in full:
        try:
            pdf.add_item(detail, client, maint)
        except Exception as e:
            print(f"  Warning: page error for {detail.get('name', '?')}: {e}")

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


def find_items_by_asset_ids(client: HomeboxClient, asset_ids: list) -> list:
    """Look up items by asset ID (e.g. '002-062') via the assets endpoint."""
    matched = []
    for raw in asset_ids:
        aid = format_asset_id(raw.strip())
        if not aid:
            continue
        item = client.get_asset(aid)
        if item:
            print(f"  Found:     [{aid}] {item.get('name')}")
            matched.append(item)
        else:
            print(f"  Not found: [{aid}]")
    return matched


def main():
    clear()
    banner()

    cfg = config.load_config()

    default_url = cfg.get("homebox_url") or "http://192.168.0.100:3100"
    url = input(f"  Homebox URL [{default_url}]: ").strip() or default_url

    saved_key = config.load_api_key()
    if saved_key:
        key = getpass.getpass("  API key [saved key found, Enter to use]: ") or saved_key
    else:
        key = getpass.getpass("  API key: ")
    if not key:
        print("\n  ERROR: An API key is required. Create one in Homebox under")
        print("  Profile > API Keys, then run this tool again.\n")
        sys.exit(1)

    print(f"\n  Connecting to {url} ...")
    client = HomeboxClient(url, key)
    ok, msg = HomeboxClient(url, key, max_retries=1).test_connection()
    if not ok:
        print(f"\n  ERROR: {msg}\n")
        sys.exit(1)

    print("  Connected!\n")

    if key != saved_key:
        try:
            config.save_api_key(key)
            print("  API key saved to Windows Credential Manager.\n")
        except Exception:
            print("  Warning: could not save the API key to Credential Manager.\n")

    default_owner = cfg.get("owner", "")
    owner_prompt = f"  Your name (for cover page) [{default_owner}]: " if default_owner else "  Your name (for cover page): "
    owner = input(owner_prompt).strip() or default_owner

    config.save_config({"homebox_url": url, "owner": owner})

    while True:
        print()
        print("  " + "-" * 42)
        print("  1)  Search items")
        print("  2)  Export by Asset ID(s)")
        print("  3)  Export ALL items")
        print("  4)  Quit")
        print("  " + "-" * 42)

        choice = input("  > ").strip()

        try:
            if choice == "1":
                q = input("\n  Search: ").strip()
                if not q:
                    continue
                print("  Searching...")
                results, truncated = client.search_items_all(q)
                display_items(results)
                if truncated:
                    print("  Note: result list was truncated.\n")
                if results:
                    ans = input(f"  Export these {len(results)} item(s) to PDF? (y/n): ").strip().lower()
                    if ans == "y":
                        generate_pdf(client, results, owner)

            elif choice == "2":
                print("\n  Enter Asset IDs separated by commas")
                print("  Example: 002-062, 002-063")
                raw = input("  Asset IDs: ").strip()
                if not raw:
                    continue
                ids = [x for x in raw.split(",") if x.strip()]
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
                items, truncated = client.search_items_all("")
                print(f"  Found {len(items)} items.")
                display_items(items[:15])
                if len(items) > 15:
                    print(f"  ... and {len(items) - 15} more\n")
                if truncated:
                    print("  Note: result list was truncated.\n")
                if items:
                    ans = input(f"  Export ALL {len(items)} items to PDF? (y/n): ").strip().lower()
                    if ans == "y":
                        generate_pdf(client, items, owner)

            elif choice == "4":
                print("\n  Goodbye!\n")
                break

            else:
                print("  Invalid choice.\n")

        except AuthError as e:
            print(f"\n  ERROR: {e}\n")
            sys.exit(1)
        except HomeboxError as e:
            print(f"\n  ERROR: {e}\n")


if __name__ == "__main__":
    main()
