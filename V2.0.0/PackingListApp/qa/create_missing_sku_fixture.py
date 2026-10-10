"""Create isolated product-review trials; do not open the app or its catalog."""
import argparse
from pathlib import Path
import sys

from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[2]
DESTINATION = ROOT / "Test Missing Product" / "20261010" / "FLASH" / "MISSING SKU TEST.pdf"


def create(with_sku=False):
    destination = DESTINATION.with_name("SKU WITHOUT STANDARD NAMES TEST.pdf") if with_sku else DESTINATION
    order_id = "TESTWITHSKU20261010B" if with_sku else "TESTNOSKU20261010A"
    product = "AURORA PEARL" if with_sku else "MOONLIGHT VANILLA"
    sku = "TEST-AURORA-50ML" if with_sku else ""
    if destination.exists():
        raise FileExistsError(f"Preserving existing fixture: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(destination), pagesize=A4, invariant=True)
    c.setTitle("Test only - SKU without standard names" if with_sku else "Test only - missing seller SKU")
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
    text(23, 74, f"Order ID: {order_id}", 9, True)
    text(23, 91, "package 1", 8)
    text(23, 110, "Parent Product Name", 8, True)

    anchors = [(23, "#"), (35, "Product"), (144, "SKU"), (169, "Variation"),
               (208, "Qty"), (230, "Unit"), (259, "Total")]
    if with_sku:
        anchors = [(23, "#"), (35, "Product"), (134, "SKU"), (189, "Variation"),
                   (225, "Qty"), (245, "Unit"), (264, "Total")]
    c.setFillColor(HexColor("#E6E9E6"))
    c.rect(18, height-136, width/2-36, 19, stroke=0, fill=1)
    for x, label in anchors:
        text(x, 130, label, 7, True)
    text(23, 153, "1", 8)
    text(35, 153, "TEST PRODUCT -", 8)
    text(35, 166, product, 8)
    # The original trial intentionally leaves SKU blank; both leave prices blank.
    if with_sku:
        text(134, 153, "TEST-AURORA-", 6.5)
        text(134, 166, "50ML", 6.5)
    text(190 if with_sku else 170, 153, "50 mL", 8)
    text(228 if with_sku else 211, 153, "1", 8)
    c.setStrokeColor(line)
    c.line(19, height-184, width/2-19, height-184)
    text(23, 205, f"Order ID: {order_id}", 8, True)
    text(23, 221, "package 1", 8)
    text(23, 247, "Test fixture: SKU present; standard names undefined." if with_sku
         else "Test fixture: one product, one piece, no seller SKU.", 7)
    text(23, 260, "No real customer, payment, or dispatch information.", 7)
    text(23, 280, "The app should request Product Name and Size.", 7)
    text(23, 294, "You may defer the prompt and generate a test list.", 7)
    c.showPage()
    c.save()

    sys.path.insert(0, str(ROOT / "PackingListApp" / "src"))
    from extractors import parse_pdf
    batch = parse_pdf(destination)
    assert batch.platform == "FLASH"
    assert len(batch.orders) == 1 and batch.pieces == 1
    item = batch.orders[0].items[0]
    assert item.product == f"TEST PRODUCT - {product}"
    assert item.variation == "50 mL" and item.pieces == 1
    assert item.seller_sku == sku
    assert "".join(item.raw_seller_sku.split()) == sku
    assert item.unit_price is None and item.total_price is None
    print(destination)
    print(batch.to_dict())


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--with-sku", action="store_true", help="Create the second trial with a SKU but no saved standard names.")
    create(parser.parse_args().with_sku)
