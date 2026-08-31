#!/usr/bin/env python3
"""Render the distinct verifier-private confirm invoice source."""
import json
from pathlib import Path
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
truth = json.loads((ROOT / "truth.json").read_text(encoding="utf-8"))
output = ROOT / "input_files" / truth["document"]["filename"]
blue = HexColor("#173A5E")
light = HexColor("#DCEAF5")

def money(value):
    return f"$ {value:,.2f}"

def page_header(pdf, page):
    pdf.setFillColor(blue)
    pdf.rect(0, 720, 612, 72, fill=1, stroke=0)
    pdf.setFillColor(white)
    pdf.setFont("Helvetica-Bold", 21)
    pdf.drawString(42, 754, truth["headers"]["vendor"])
    pdf.setFont("Helvetica", 10)
    pdf.drawRightString(570, 754, f"INVOICE {truth['headers']['invoice_id']}")
    pdf.drawRightString(570, 738, f"Page {page} of 2")
    pdf.setFillColor(black)

pdf = canvas.Canvas(str(output), pagesize=letter)
pdf.setTitle(f"Confirm invoice {truth['headers']['invoice_id']}")
page_header(pdf, 1)
pdf.setFont("Helvetica-Bold", 11)
pdf.drawString(42, 690, "Bill to")
pdf.setFont("Helvetica", 11)
pdf.drawString(42, 673, truth["headers"]["customer"])
pdf.drawString(42, 656, f"Currency: {truth['headers']['currency']}")
pdf.drawRightString(570, 690, f"Invoice date: {truth['headers']['invoice_date']}")
pdf.drawRightString(570, 673, f"Purchase order: {truth['headers']['po_reference']}")

headers = ["Line", "Description", "Qty", "Unit price", "Line total"]
x = [42, 92, 328, 390, 480]
pdf.setFillColor(light)
pdf.rect(42, 610, 528, 26, fill=1, stroke=0)
pdf.setFillColor(black)
pdf.setFont("Helvetica-Bold", 10)
for pos, label in zip(x, headers):
    pdf.drawString(pos, 619, label)
pdf.setFont("Helvetica", 10)
y = 584
for row in truth["line_items"][:3]:
    pdf.drawString(x[0], y, row["line_id"])
    pdf.drawString(x[1], y, row["description"])
    pdf.drawRightString(370, y, str(row["quantity"]))
    pdf.drawRightString(465, y, money(row["unit_price"]))
    pdf.drawRightString(570, y, money(row["quantity"] * row["unit_price"]))
    pdf.line(42, y - 8, 570, y - 8)
    y -= 36
pdf.setFont("Helvetica-Bold", 11)
pdf.drawString(42, 438, "Continuation")
pdf.setFont("Helvetica", 10)
pdf.drawString(42, 420, "The calibration dossier and invoice summary continue on page 2.")
pdf.drawString(42, 398, "Payment terms and all amounts are stated in USD.")
pdf.setFont("Helvetica-Oblique", 9)
pdf.drawString(42, 56, "Confirm sibling source - distinct from the development invoice instance.")
pdf.showPage()

page_header(pdf, 2)
row = truth["line_items"][3]
pdf.setFillColor(light)
pdf.rect(42, 650, 528, 26, fill=1, stroke=0)
pdf.setFillColor(black)
pdf.setFont("Helvetica-Bold", 10)
for pos, label in zip(x, headers):
    pdf.drawString(pos, 659, label)
pdf.setFont("Helvetica", 10)
pdf.drawString(x[0], 624, row["line_id"])
pdf.drawString(x[1], 624, row["description"])
pdf.drawRightString(370, 624, str(row["quantity"]))
pdf.drawRightString(465, 624, money(row["unit_price"]))
pdf.drawRightString(570, 624, money(row["quantity"] * row["unit_price"]))
subtotal = sum(item["quantity"] * item["unit_price"] for item in truth["line_items"])
discount = subtotal * truth["terms"]["discount_rate"]
taxable = subtotal - discount
tax = taxable * truth["terms"]["tax_rate"]
total = taxable + tax + truth["terms"]["freight"]
summary = [
    ("Subtotal", subtotal),
    (f"Discount ({truth['terms']['discount_rate']:.0%})", -discount),
    ("Taxable amount", taxable),
    (f"Sales tax ({truth['terms']['tax_rate']:.2%})", tax),
    ("Freight", truth["terms"]["freight"]),
    ("Amount payable", total),
]
pdf.setFont("Helvetica", 11)
y = 548
for label, value in summary:
    if label == "Amount payable":
        pdf.setFont("Helvetica-Bold", 12)
        pdf.setFillColor(blue)
    pdf.drawString(330, y, label)
    pdf.drawRightString(570, y, money(value))
    pdf.setFillColor(black)
    if label == "Amount payable":
        pdf.setFont("Helvetica", 11)
    y -= 30
pdf.setFillColor(light)
pdf.rect(42, 290, 528, 70, fill=1, stroke=0)
pdf.setFillColor(black)
pdf.setFont("Helvetica-Bold", 10)
pdf.drawString(56, 336, "Remittance note")
pdf.setFont("Helvetica", 10)
pdf.drawString(56, 316, f"Reference invoice {truth['headers']['invoice_id']} and PO {truth['headers']['po_reference']}.")
pdf.drawString(56, 298, f"Customer: {truth['headers']['customer']} | Currency: {truth['headers']['currency']}")
pdf.save()
print(output)
