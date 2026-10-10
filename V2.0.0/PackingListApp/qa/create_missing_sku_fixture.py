"""Create the isolated missing-SKU user trial; do not open the app or its catalog."""
from pathlib import Path
import sys

from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[2]
DESTINATION = ROOT / "Test Missing Product" / "20261010" / "FLASH" / "MISSING SKU TEST.pdf"


def create():
    if DESTINATION.exists():
        raise FileExistsError(f"Preserving existing fixture: {DESTINATION}")
    DESTINATION.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(DESTINATION), pagesize=A4, invariant=True)
    c.setTitle("Test only - missing seller SKU")
    c.setAuthor("Packing List Manager test fixture")
    width, height = A4
    ink = HexColor("#24292C")
    paper = HexColor("#F1F2EF")
    line = HexColor("#6B767C")

    def text(x, y, value, size=7, bold=False):
        c.setFillColor(ink)
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(x, height-y, value)

    # The supported Shopee packing-list format is four labels on an A4 page.
    # Populate only the top-left label; leave the other three quadrants empty.
    c.setFillColor(paper)
    c.setStrokeColor(line)
    c.roundRect(13, height-315, width/2-26, 301, 3, stroke=1, fill=1)
    text(23, 35, "TEST ONLY - DO NOT SHIP", 12, True)
    text(23, 53, "Synthetic Shopee-format packing list", 8)
    text(23, 74, "Order ID: TESTNOSKU20261010A", 9, True)
    text(23, 91, "package 1", 8)
    text(23, 110, "Parent Product Name", 8, True)

    anchors = [(23, "#"), (35, "Product"), (144, "SKU"), (169, "Variation"),
               (208, "Qty"), (230, "Unit"), (259, "Total")]
    c.setFillColor(HexColor("#E6E9E6"))
    c.rect(18, height-136, width/2-36, 19, stroke=0, fill=1)
    for x, label in anchors:
        text(x, 130, label, 7, True)
    text(23, 153, "1", 8)
    text(35, 153, "TEST PRODUCT -", 8)
    text(35, 166, "MOONLIGHT VANILLA", 8)
    # Intentionally blank SKU and prices. Do not put a placeholder in the cell.
    text(170, 153, "50 mL", 8)
    text(211, 153, "1", 8)
    c.setStrokeColor(line)
    c.line(19, height-184, width/2-19, height-184)
    text(23, 205, "Order ID: TESTNOSKU20261010A", 8, True)
    text(23, 221, "package 1", 8)
    text(23, 247, "Test fixture: one product, one piece, no seller SKU.", 7)
    text(23, 260, "No real customer, payment, or dispatch information.", 7)
    text(23, 280, "The app should request Product Name and Size.", 7)
    text(23, 294, "You may defer the prompt and generate a test list.", 7)
    c.showPage()
    c.save()

    sys.path.insert(0, str(ROOT / "PackingListApp" / "src"))
    from extractors import parse_pdf
    batch = parse_pdf(DESTINATION)
    assert batch.platform == "FLASH"
    assert len(batch.orders) == 1 and batch.pieces == 1
    item = batch.orders[0].items[0]
    assert item.product == "TEST PRODUCT - MOONLIGHT VANILLA"
    assert item.variation == "50 mL" and item.pieces == 1
    assert item.seller_sku == item.raw_seller_sku == ""
    assert item.unit_price is None and item.total_price is None
    print(DESTINATION)
    print(batch.to_dict())


if __name__ == "__main__":
    create()
