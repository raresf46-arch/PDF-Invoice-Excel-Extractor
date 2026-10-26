"""Generate demo invoice PDFs in two different supplier layouts (EN and RO)."""
import random
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

random.seed(11)
OUT = Path(__file__).parent / "sample_invoices"
OUT.mkdir(exist_ok=True)
styles = getSampleStyleSheet()
ITEMS = [("Sensor Kit", 129.90), ("Cable Set 5m", 24.50), ("Control Unit", 410.00),
         ("Mounting Kit", 38.75), ("Service Hours", 55.00), ("Pressure Valve", 87.20)]


def line_items():
    picks = random.sample(ITEMS, random.randint(2, 5))
    return [(name, random.randint(1, 12), price) for name, price in picks]


def layout_en(path, inv_no, date, supplier, items, wrong_total=False):
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    subtotal = sum(q * p for _, q, p in items)
    vat = round(subtotal * 0.19, 2)
    total = round(subtotal + vat + (50 if wrong_total else 0), 2)
    data = [["Description", "Qty", "Unit Price", "Amount"]] + \
           [[n, str(q), f"{p:,.2f}", f"{q * p:,.2f}"] for n, q, p in items]
    t = Table(data, colWidths=[230, 50, 90, 90])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                           ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                           ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                           ("ALIGN", (1, 1), (-1, -1), "RIGHT")]))
    doc.build([Paragraph(f"<b>{supplier}</b>", styles["Title"]),
               Paragraph("INVOICE", styles["Heading2"]),
               Paragraph(f"Invoice No: {inv_no}", styles["Normal"]),
               Paragraph(f"Invoice Date: {date}", styles["Normal"]),
               Paragraph("Bill To: Demo Client SRL, Craiova", styles["Normal"]),
               Spacer(1, 18), t, Spacer(1, 14),
               Paragraph(f"Subtotal: {subtotal:,.2f} EUR", styles["Normal"]),
               Paragraph(f"VAT 19%: {vat:,.2f} EUR", styles["Normal"]),
               Paragraph(f"<b>Total Due: {total:,.2f} EUR</b>", styles["Normal"])])


def ro_num(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def layout_ro(path, inv_no, date, supplier, items):
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    subtotal = sum(q * p for _, q, p in items)
    vat = round(subtotal * 0.19, 2)
    total = round(subtotal + vat, 2)
    data = [["Nr.", "Denumire produs", "Cant.", "Pret unitar", "Valoare"]] + \
           [[str(i + 1), n, str(q), ro_num(p), ro_num(q * p)] for i, (n, q, p) in enumerate(items)]
    t = Table(data, colWidths=[30, 210, 50, 90, 90])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF4")),
                           ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                           ("ALIGN", (2, 1), (-1, -1), "RIGHT")]))
    doc.build([Paragraph(f"Furnizor: <b>{supplier}</b>", styles["Normal"]),
               Paragraph("CUI: RO00000000 (demo)", styles["Normal"]), Spacer(1, 10),
               Paragraph("FACTURA FISCALA", styles["Title"]),
               Paragraph(f"Factura nr. {inv_no} / Data emiterii: {date}", styles["Normal"]),
               Spacer(1, 18), t, Spacer(1, 14),
               Paragraph(f"Total fara TVA: {ro_num(subtotal)} RON", styles["Normal"]),
               Paragraph(f"TVA 19%: {ro_num(vat)} RON", styles["Normal"]),
               Paragraph(f"<b>Total de plata: {ro_num(total)} RON</b>", styles["Normal"])])


for i in range(5):
    layout_en(OUT / f"invoice_NTS-{2401 + i}.pdf", f"NTS-{2401 + i}", f"2026-0{i + 1}-1{i}",
              "Northline Tech Supplies", line_items(), wrong_total=(i == 3))
for i in range(4):
    layout_ro(OUT / f"factura_DMC{780 + i}.pdf", f"DMC{780 + i}", f"{12 + i}.0{i + 2}.2026",
              "Danubia Components SRL", line_items())
print("Invoices written to", OUT)
