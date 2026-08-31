#!/usr/bin/env python3
"""Render the three distinct verifier-private confirm receipt images."""
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
truth = json.loads((ROOT / "truth.json").read_text(encoding="utf-8"))
font_path = "/System/Library/Fonts/Supplemental/Arial.ttf"
bold_path = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
title_font = ImageFont.truetype(bold_path, 42)
body_font = ImageFont.truetype(font_path, 27)
bold_font = ImageFont.truetype(bold_path, 27)
small_font = ImageFont.truetype(font_path, 21)

for document in truth["documents"]:
    image = Image.new("RGB", (1000, 1400), "#FBFAF5")
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 1000, 160), fill="#2D4D4A")
    draw.text((60, 48), document["merchant"], font=title_font, fill="white")
    draw.text((60, 205), f"Receipt date: {document['date']}", font=body_font, fill="#263331")
    draw.text((60, 250), f"Receipt file: {document['filename']}", font=small_font, fill="#536562")
    draw.line((60, 305, 940, 305), fill="#879A96", width=3)
    rows = [row for row in truth["items"] if row["document_id"] == document["document_id"]]
    y = 355
    for row in rows:
        draw.text((70, y), row["item"], font=body_font, fill="#263331")
        draw.text((750, y), f"$ {row['amount']:,.2f}", font=body_font, fill="#263331")
        draw.text((70, y + 40), f"Category: {row['category']} | Modifier: {row['modifier']}", font=small_font, fill="#60706D")
        draw.line((60, y + 85, 940, y + 85), fill="#CDD5D2", width=2)
        y += 120
    subtotal = sum(row["amount"] for row in rows)
    tax = subtotal * document["tax_rate"]
    total = subtotal + tax + document["tip"]
    y = max(y + 20, 730)
    for label, value in [
        ("Subtotal", subtotal),
        (f"Tax ({document['tax_rate']:.2%})", tax),
        ("Tip", document["tip"]),
        ("Total", total),
    ]:
        font = bold_font if label == "Total" else body_font
        draw.text((510, y), label, font=font, fill="#263331")
        draw.text((750, y), f"$ {value:,.2f}", font=font, fill="#263331")
        y += 58
    draw.rectangle((50, 1120, 950, 1240), outline="#9BAAA7", width=3)
    draw.text((75, 1150), "Item descriptions, tax rate, tip, and amounts are", font=small_font, fill="#536562")
    draw.text((75, 1185), "visible source facts for the confirm receipt batch.", font=small_font, fill="#536562")
    image.save(ROOT / "input_files" / document["filename"], quality=94)
    print(ROOT / "input_files" / document["filename"])
