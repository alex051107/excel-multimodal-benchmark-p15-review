#!/usr/bin/env python3
"""Render the distinct verifier-private confirm bank statement source."""
import json
from pathlib import Path
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
truth = json.loads((ROOT / "truth.json").read_text(encoding="utf-8"))
output = ROOT / "input_files" / truth["document"]["filename"]
green = HexColor("#1F5A4A")
light = HexColor("#E2F0EB")

def money(value):
    return f"$ {value:,.2f}"

def header(pdf, page):
    pdf.setFillColor(green)
    pdf.rect(0, 724, 612, 68, fill=1, stroke=0)
    pdf.setFillColor(white)
    pdf.setFont("Helvetica-Bold", 20)
    pdf.drawString(42, 756, truth["headers"]["bank"])
    pdf.setFont("Helvetica", 10)
    pdf.drawRightString(570, 756, f"Account ending {truth['headers']['account_suffix']}")
    pdf.drawRightString(570, 740, f"Page {page} of 2")
    pdf.setFillColor(black)

pdf = canvas.Canvas(str(output), pagesize=letter)
pdf.setTitle("Confirm bank statement")
header(pdf, 1)
pdf.setFont("Helvetica", 11)
pdf.drawString(42, 690, f"Statement period: {truth['headers']['period_start']} through {truth['headers']['period_end']}")
pdf.drawString(42, 670, f"Currency: {truth['headers']['currency']}")
pdf.setFillColor(light)
pdf.rect(42, 624, 528, 28, fill=1, stroke=0)
pdf.setFillColor(black)
pdf.setFont("Helvetica-Bold", 11)
pdf.drawString(54, 634, f"Opening balance: {money(truth['headers']['opening_balance'])}")
columns = [42, 128, 420, 510, 548]
pdf.setFont("Helvetica-Bold", 9)
pdf.drawString(columns[0], 590, "Date")
pdf.drawString(columns[1], 590, "Description")
pdf.drawRightString(columns[2], 590, "Debit")
pdf.drawRightString(columns[3], 590, "Credit")
pdf.drawString(columns[4], 590, "Page")
y = 556
for row in truth["transactions"][:3]:
    pdf.setFont("Helvetica", 9)
    pdf.drawString(columns[0], y, row["date"])
    pdf.drawString(columns[1], y, row["description"])
    pdf.drawRightString(columns[2], y, money(row["debit"]) if row["debit"] else "-")
    pdf.drawRightString(columns[3], y, money(row["credit"]) if row["credit"] else "-")
    pdf.drawString(columns[4], y, str(row["page"]))
    pdf.line(42, y - 8, 570, y - 8)
    y -= 38
pdf.setFillColor(light)
pdf.rect(42, 402, 528, 58, fill=1, stroke=0)
pdf.setFillColor(black)
pdf.setFont("Helvetica-Bold", 10)
pdf.drawString(56, 438, "Continuation notice")
pdf.setFont("Helvetica", 10)
pdf.drawString(56, 418, "Additional debit activity and the statement closing balance appear on page 2.")
pdf.showPage()

header(pdf, 2)
pdf.setFont("Helvetica-Bold", 9)
pdf.drawString(columns[0], 664, "Date")
pdf.drawString(columns[1], 664, "Description")
pdf.drawRightString(columns[2], 664, "Debit")
pdf.drawRightString(columns[3], 664, "Credit")
pdf.drawString(columns[4], 664, "Page")
y = 632
for row in truth["transactions"][3:]:
    pdf.setFont("Helvetica", 9)
    pdf.drawString(columns[0], y, row["date"])
    pdf.drawString(columns[1], y, row["description"])
    pdf.drawRightString(columns[2], y, money(row["debit"]) if row["debit"] else "-")
    pdf.drawRightString(columns[3], y, money(row["credit"]) if row["credit"] else "-")
    pdf.drawString(columns[4], y, str(row["page"]))
    pdf.line(42, y - 8, 570, y - 8)
    y -= 38
credits = sum(row["credit"] for row in truth["transactions"])
debits = sum(row["debit"] for row in truth["transactions"])
closing = truth["headers"]["opening_balance"] + credits - debits
pdf.setFillColor(light)
pdf.rect(300, 420, 270, 126, fill=1, stroke=0)
pdf.setFillColor(black)
pdf.setFont("Helvetica", 11)
pdf.drawString(316, 518, "Opening balance")
pdf.drawRightString(554, 518, money(truth["headers"]["opening_balance"]))
pdf.drawString(316, 492, "Total credits")
pdf.drawRightString(554, 492, money(credits))
pdf.drawString(316, 466, "Total debits")
pdf.drawRightString(554, 466, money(debits))
pdf.setFont("Helvetica-Bold", 12)
pdf.setFillColor(green)
pdf.drawString(316, 438, "Closing balance")
pdf.drawRightString(554, 438, money(closing))
pdf.setFillColor(black)
pdf.setFont("Helvetica", 9)
pdf.drawString(42, 360, "Transaction descriptions and page continuation are part of this statement record.")
pdf.save()
print(output)
