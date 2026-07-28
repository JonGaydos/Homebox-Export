from hbx.report import InventoryReport, fmt_date, fmt_price, loc_name, sanitize_text


def test_sanitize_text_replaces_smart_punctuation():
    assert sanitize_text("‘a’ “b” – c…") == "'a' \"b\" - c..."
    assert sanitize_text(None) == ""


def test_fmt_date_handles_iso_and_zero_year():
    assert fmt_date("2026-07-19T22:52:00.875242974Z") == "Jul 19, 2026"
    assert fmt_date("2024-09-07") == "Sep 07, 2024"
    assert fmt_date("0001-01-01T00:00:00Z") == ""
    assert fmt_date("") == ""
    assert fmt_date(None) == ""


def test_fmt_price():
    assert fmt_price(1234.5) == "$1,234.50"
    assert fmt_price(0) == ""
    assert fmt_price(None) == ""
    assert fmt_price("junk") == ""


def test_loc_name_reads_parent():
    assert loc_name({"parent": {"name": "Garage"}}) == "Garage"
    assert loc_name({"parent": None}) == ""
    assert loc_name({}) == ""


class NoImageClient:
    def get_attachment(self, entity_id, attachment_id):
        raise RuntimeError("no images in tests")


def make_item(**overrides):
    item = {
        "id": "e1",
        "assetId": "002-062",
        "name": "Test Widget",
        "description": "A widget.",
        "quantity": 2,
        "insured": True,
        "purchasePrice": 19.99,
        "purchaseDate": "2026-01-15T00:00:00Z",
        "purchaseFrom": "Acme",
        "serialNumber": "SN1",
        "modelNumber": "M1",
        "manufacturer": "Acme Corp",
        "lifetimeWarranty": False,
        "warrantyExpires": "",
        "warrantyDetails": "",
        "soldDate": "",
        "soldTo": "",
        "soldPrice": 0,
        "soldNotes": "",
        "notes": "",
        "parent": {"name": "Garage"},
        "tags": [{"name": "tools"}],
        "attachments": [],
        "fields": [],
    }
    item.update(overrides)
    return item


def test_cover_page_has_no_footer(tmp_path):
    pypdf = __import__("pytest").importorskip("pypdf")
    pdf = InventoryReport()
    pdf.add_cover("Jon")
    pdf.add_item(make_item(), NoImageClient(), maintenance=[])
    out = tmp_path / "cover.pdf"
    pdf.output(str(out))
    pages = pypdf.PdfReader(str(out)).pages
    assert "Page 0" not in pages[0].extract_text()
    assert "Page 1" in pages[1].extract_text()


def test_report_generates_pdf_bytes(tmp_path):
    pdf = InventoryReport()
    pdf.add_cover("Jon")
    items = [make_item(), make_item(assetId="002-063", name="Second")]
    pdf.add_summary(items)
    for it in items:
        pdf.add_item(it, NoImageClient(), maintenance=[])
    out = tmp_path / "report.pdf"
    pdf.output(str(out))
    data = out.read_bytes()
    assert data.startswith(b"%PDF")
    assert pdf.item_count == 2
    assert pdf.total_value == 19.99 * 2
