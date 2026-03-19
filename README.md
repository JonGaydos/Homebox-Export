# Homebox Export

A standalone Windows application that connects to your [Homebox](https://github.com/sysadminsmedia/homebox) instance and generates professional PDF inventory reports — designed for insurance documentation, warranty claims, and asset records.

![License](https://img.shields.io/github/license/JonGaydos/Homebox-Export)
![Release](https://img.shields.io/github/v/release/JonGaydos/Homebox-Export)

## Features

- **Search & browse** your Homebox inventory from a native Windows GUI
- **Look up items** by asset ID (single or comma-separated)
- **Export to PDF** — single item or full inventory
- **Professional insurance-grade layout** with cover page, summary table, and per-item detail pages
- **Embeds photos and receipt images** directly in the PDF (skips manuals and other documents)
- **Sortable item table** — click column headers to sort
- **Portable config** — settings file lives next to the .exe, move it wherever you want

### What's in the PDF?

Each item page includes:

- Primary photo + additional photos
- Asset ID, serial/model number, manufacturer
- Location, tags, quantity, insured status
- Purchase info (from, date, price)
- Warranty details
- Custom fields and notes
- Receipt images (embedded)
- Maintenance history
- Sold information (if applicable)

Multi-item exports include a **cover page** and **summary page** with total item count, estimated value, and insured item count.

## Download

Download the latest `HomeboxExport.exe` from the [Releases](https://github.com/JonGaydos/Homebox-Export/releases) page. No installation or Python required — just run it.

## Usage

1. Run `HomeboxExport.exe`
2. Enter your Homebox URL (e.g., `http://192.168.1.50:7745`), username, and password
3. Click **Connect**
4. Search for items or click **Load All**
5. Select items and click **Export Selected to PDF**, or export everything with **Export All to PDF**
6. Choose where to save and you're done

Your connection settings (URL, username, display name) are saved to `homebox_export_config.json` next to the .exe for convenience. **Passwords are never saved.**

## Building from Source

### Requirements

- Python 3.10+
- pip

### Run directly

```bash
pip install -r requirements.txt
python homebox_export_gui.py
```

### Build the .exe

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name "HomeboxExport" --collect-data fpdf2 homebox_export_gui.py
```

The .exe will be in the `dist/` folder.

There is also a CLI version available (`homebox_export.py`) if you prefer terminal-based usage.

## Homebox API

This tool uses the Homebox REST API and requires no changes to your Homebox instance. It authenticates with your existing credentials and reads item data, maintenance logs, and attachment images through the standard API endpoints.

## Inspired by

[Homebox Discussion #735](https://github.com/sysadminsmedia/homebox/discussions/735) — "Print item for insurance/legal"

## License

MIT
