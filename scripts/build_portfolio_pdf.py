#!/usr/bin/env python3
"""Build the compact Coco You portfolio PDF."""

from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image, ImageDraw
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "pdf" / "coco-you-portfolio.pdf"
ASSET = ROOT / "assets" / "coco-you-portfolio.pdf"
PORTRAIT = ROOT / "assets" / "headshot-hd.png"
TMP_DIR = ROOT / "tmp" / "pdfs"
PORTRAIT_ROUND = TMP_DIR / "portfolio-headshot-round.png"
COVER = TMP_DIR / "portfolio-cover.pdf"
ORIGINAL = TMP_DIR / "original-coco-you-portfolio.pdf"


def hex_color(value: str) -> colors.Color:
    return colors.HexColor(value)


def make_round_portrait() -> Path:
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    img = Image.open(PORTRAIT).convert("RGB")
    crop_w = min(img.width, int(img.height * 0.76))
    left = (img.width - crop_w) // 2
    top = int(img.height * 0.05)
    cropped = img.crop((left, top, left + crop_w, top + crop_w))
    size = 720
    cropped = cropped.resize((size, size), Image.Resampling.LANCZOS)

    mask = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, size, size), radius=90, fill=255)

    out = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    out.paste(cropped, (0, 0), mask)
    out.save(PORTRAIT_ROUND)
    return PORTRAIT_ROUND


def draw_wrapped(c: canvas.Canvas, text: str, x: float, y: float, width: float, font: str, size: float, leading: float, color: str) -> float:
    c.setFont(font, size)
    c.setFillColor(hex_color(color))
    words = text.split()
    line = ""
    for word in words:
        trial = f"{line} {word}".strip()
        if c.stringWidth(trial, font, size) <= width:
            line = trial
        else:
            c.drawString(x, y, line)
            y -= leading
            line = word
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y


def draw_chip(c: canvas.Canvas, text: str, x: float, y: float) -> float:
    font = "Helvetica-Bold"
    size = 8.6
    pad_x = 10
    w = c.stringWidth(text, font, size) + pad_x * 2
    h = 21
    c.setFillColor(hex_color("#fff7fa"))
    c.setStrokeColor(hex_color("#e7c4ce"))
    c.setLineWidth(0.7)
    c.roundRect(x, y - h, w, h, 10, fill=1, stroke=1)
    c.setFillColor(hex_color("#6a2436"))
    c.setFont(font, size)
    c.drawCentredString(x + w / 2, y - 14.2, text)
    return w


def draw_section_box(c: canvas.Canvas, title: str, x: float, y: float, w: float, h: float) -> None:
    c.setFillColor(colors.white)
    c.setStrokeColor(hex_color("#ead1da"))
    c.setLineWidth(0.9)
    c.roundRect(x, y - h, w, h, 10, fill=1, stroke=1)
    c.setFillColor(hex_color("#6a2436"))
    c.setFont("Helvetica-Bold", 10.2)
    c.drawString(x + 18, y - 26, title.upper())
    c.setStrokeColor(hex_color("#ead1da"))
    c.line(x, y - 42, x + w, y - 42)


def build_cover() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    portrait = make_round_portrait()
    c = canvas.Canvas(str(COVER), pagesize=letter)
    page_w, page_h = letter

    bg = "#fff9fb"
    ink = "#2d2228"
    muted = "#6f5864"
    cornell = "#6a2436"
    blush = "#f3c6d2"
    pale = "#faedf5"

    c.setFillColor(hex_color(bg))
    c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

    # Left identity panel.
    panel_x, panel_y, panel_w, panel_h = 42, 48, 190, 696
    c.setFillColor(hex_color(cornell))
    c.roundRect(panel_x, panel_y, panel_w, panel_h, 18, fill=1, stroke=0)
    c.setFillColor(hex_color(blush))
    c.circle(panel_x + panel_w - 24, panel_y + panel_h - 90, 62, fill=1, stroke=0)
    c.setFillColor(hex_color("#efb0c0"))
    c.circle(panel_x + 22, panel_y + 90, 78, fill=1, stroke=0)

    c.drawImage(ImageReader(str(portrait)), panel_x + 32, panel_y + panel_h - 220, width=126, height=126, mask="auto")

    c.setFillColor(colors.white)
    c.setFont("Times-BoldItalic", 32)
    c.drawString(panel_x + 28, panel_y + panel_h - 305, "Coco")
    c.drawString(panel_x + 28, panel_y + panel_h - 346, "You")

    c.setFont("Helvetica", 11.2)
    c.drawString(panel_x + 28, panel_y + panel_h - 394, "Cornell University")
    c.drawString(panel_x + 28, panel_y + panel_h - 413, "Class of 2028")

    c.setFont("Helvetica-Bold", 8.8)
    c.setFillColor(hex_color("#f6d8df"))
    c.drawString(panel_x + 28, panel_y + 262, "PORTFOLIO")
    c.setFillColor(colors.white)
    draw_wrapped(
        c,
        "Education and skills snapshot for recruiting, research, and creative work.",
        panel_x + 28,
        panel_y + 242,
        128,
        "Helvetica",
        10.4,
        15,
        "#ffffff",
    )
    c.setFont("Helvetica", 9.8)
    c.drawString(panel_x + 28, panel_y + 72, "ly458@cornell.edu")
    c.drawString(panel_x + 28, panel_y + 54, "linkedin.com/in/cocoyou")
    c.drawString(panel_x + 28, panel_y + 36, "cocoylh.com")

    # Right content.
    x = 266
    right_w = 300
    y = 710
    c.setFillColor(hex_color(ink))
    c.setFont("Times-Bold", 25)
    c.drawString(x, y, "Portfolio Snapshot")
    y -= 28
    y = draw_wrapped(
        c,
        "Cornell student building across data, finance, product analytics, and visual storytelling.",
        x,
        y,
        right_w,
        "Helvetica",
        10.5,
        15,
        muted,
    )

    y -= 20
    draw_section_box(c, "Education", x, y, right_w, 154)
    c.setFillColor(hex_color(ink))
    c.setFont("Helvetica-Bold", 13)
    c.drawString(x + 18, y - 66, "Cornell University")
    c.setFont("Helvetica", 10.2)
    c.setFillColor(hex_color(muted))
    c.drawString(x + 18, y - 86, "B.S. in Information Science")
    c.drawString(x + 18, y - 103, "Environment and Sustainability")
    c.drawString(x + 18, y - 120, "Expected graduation: May 2028")
    c.drawString(x + 18, y - 137, "Concentration: Data Science | Minor: Business")

    y -= 182
    draw_section_box(c, "Skills", x, y, right_w, 236)
    chips = [
        "Python", "Java", "R", "JavaScript",
        "Excel", "Financial Modeling", "Data Analysis",
        "Dashboarding", "Research Writing", "Photography",
        "Photoshop", "Lightroom", "Canva",
    ]
    cx = x + 18
    cy = y - 66
    max_x = x + right_w - 18
    for chip_text in chips:
        chip_w = c.stringWidth(chip_text, "Helvetica-Bold", 8.6) + 20
        if cx + chip_w > max_x:
            cx = x + 18
            cy -= 31
        draw_chip(c, chip_text, cx, cy)
        cx += chip_w + 8

    c.setFillColor(hex_color(muted))
    c.setFont("Helvetica", 9.6)
    c.drawString(x, 70, "Cornell University | Information Science | Class of 2028")

    c.showPage()
    c.save()


def merge_with_selected_works() -> None:
    if not ORIGINAL.exists():
        raise FileNotFoundError(
            f"Missing source works PDF: {ORIGINAL}. "
            "Restore the original portfolio there before rebuilding."
        )

    cover_reader = PdfReader(str(COVER))
    works_reader = PdfReader(str(ORIGINAL))
    writer = PdfWriter()
    writer.add_page(cover_reader.pages[0])

    # Skip the old bio/education page and preserve the artwork pages.
    for page in works_reader.pages[1:]:
        writer.add_page(page)

    with OUT.open("wb") as f:
        writer.write(f)
    shutil.copyfile(OUT, ASSET)
    print(OUT)
    print(ASSET)


def build() -> None:
    build_cover()
    merge_with_selected_works()


if __name__ == "__main__":
    build()
