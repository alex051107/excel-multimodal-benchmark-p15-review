#!/usr/bin/env python3
"""Render the distinct verifier-private confirm quote scan."""
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
truth = json.loads((ROOT / "truth.json").read_text(encoding="utf-8"))
output = ROOT / "input_files" / truth["document"]["filename"]
font_path = "/System/Library/Fonts/Supplemental/Arial.ttf"
bold_path = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
title = ImageFont.truetype(bold_path, 48)
heading = ImageFont.truetype(bold_path, 28)
body = ImageFont.truetype(font_path, 25)
small = ImageFont.truetype(font_path, 21)

image = Image.new("RGB", (1400, 1700), "#F7F3EA")
draw = ImageDraw.Draw(image)
draw.rectangle((0, 0, 1400, 190), fill="#263D55")
draw.text((70, 55), truth["headers"]["vendor"], font=title, fill="white")
draw.text((1000, 66), "CONTRACTOR QUOTE", font=heading, fill="#DDE8F2")
draw.text((70, 235), f"Quote ID: {truth['headers']['quote_id']}", font=heading, fill="#1E2D3D")
draw.text((70, 285), f"Customer: {truth['headers']['customer']}", font=body, fill="#1E2D3D")
draw.text((70, 330), f"Issue date: {truth['headers']['quote_date']}", font=body, fill="#1E2D3D")
draw.text((760, 330), f"Valid through: {truth['headers']['valid_through']}", font=body, fill="#1E2D3D")
draw.text((70, 375), f"Currency: {truth['headers']['currency']}", font=body, fill="#1E2D3D")

draw.rectangle((60, 450, 1340, 510), fill="#D7E3EE")
for x, label in [(80, "Scope"), (330, "Description"), (1020, "Amount")]:
    draw.text((x, 466), label, font=heading, fill="#1E2D3D")
y = 545
for row in truth["line_items"]:
    draw.text((80, y), row["group"], font=body, fill="#1E2D3D")
    draw.text((330, y), row["description"], font=body, fill="#1E2D3D")
    draw.text((1020, y), f"$ {row['amount']:,.2f}", font=body, fill="#1E2D3D")
    draw.line((70, y + 43, 1330, y + 43), fill="#AAB8C5", width=2)
    y += 92

base = sum(row["amount"] for row in truth["line_items"] if not row["optional"])
discount = base * truth["terms"]["discount_rate"]
taxable = base - discount
tax = taxable * truth["terms"]["tax_rate"]
total = taxable + tax
draw.rectangle((650, 890, 1340, 1190), fill="#E9EFF4")
summary = [
    ("Base-scope subtotal", base),
    (f"Discount ({truth['terms']['discount_rate']:.0%})", -discount),
    (f"Tax ({truth['terms']['tax_rate']:.2%})", tax),
    ("Base-scope total", total),
]
y = 930
for label, value in summary:
    font = heading if label == "Base-scope total" else body
    draw.text((690, y), label, font=font, fill="#1E2D3D")
    draw.text((1100, y), f"$ {value:,.2f}", font=font, fill="#1E2D3D")
    y += 65
draw.rectangle((60, 1260, 1340, 1435), outline="#8EA2B5", width=3)
draw.text((85, 1290), "Optional alternate treatment", font=heading, fill="#1E2D3D")
draw.text((85, 1345), "The Two-year service plan is retained separately and excluded", font=body, fill="#1E2D3D")
draw.text((85, 1385), "from the base-scope subtotal, discount, tax, and total.", font=body, fill="#1E2D3D")
draw.text((70, 1600), "Confirm sibling source - distinct from the development quote instance.", font=small, fill="#53687A")
image.save(output, quality=94)
print(output)
