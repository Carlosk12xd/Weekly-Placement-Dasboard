"""Branding functions adapted from Carlosk12xd/Excel-to-PDF, commit 471bf7174d44a0ab08439128d4853096d0fbc693."""

from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFont

PAGE_W = 1920

PAGE_H = 1080

SIDEBAR_W = 170

MARGIN_X = 20

MARGIN_Y = 20

CONTENT_W = PAGE_W - SIDEBAR_W - 2 * MARGIN_X - 10

CONTENT_H = PAGE_H - 2 * MARGIN_Y

SIDEBAR_BG = (245, 246, 248)

DIVIDER = (205, 208, 214)

TITLE_COLOR = (16, 49, 101)

TEXT_COLOR = (45, 45, 45)

def crop_white_space(image_path: Path, output_path: Path, threshold: int = 12, margin_px: int = 18) -> Path:
    image = Image.open(image_path).convert("RGB")
    white = Image.new("RGB", image.size, (255, 255, 255))
    diff = ImageChops.difference(image, white)
    mask = diff.point(lambda p: 255 if p > threshold else 0)
    bbox = mask.getbbox()
    if bbox is None:
        image.save(output_path)
        return output_path
    left = max(0, bbox[0] - margin_px)
    top = max(0, bbox[1] - margin_px)
    right = min(image.width, bbox[2] + margin_px)
    bottom = min(image.height, bbox[3] + margin_px)
    image.crop((left, top, right, bottom)).save(output_path)
    return output_path

def resize_for_page(image_path: Path, output_path: Path, max_dimension: int = 2600) -> Path:
    image = Image.open(image_path).convert("RGB")
    scale = min(max_dimension / image.width, max_dimension / image.height, 1.0)
    if scale < 1.0:
        image = image.resize((int(image.width * scale), int(image.height * scale)), Image.LANCZOS)
    image.save(output_path, quality=94, optimize=True)
    return output_path

def transparent_circle_logo(source: Path, output: Path) -> Path:
    image = Image.open(source).convert("RGBA")
    mask = Image.new("L", image.size, 0)
    draw = ImageDraw.Draw(mask)
    pad = int(min(image.size) * 0.01)
    draw.ellipse((pad, pad, image.width - pad, image.height - pad), fill=255)
    image.putalpha(mask)
    image.save(output)
    return output

def load_font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()

def fit_text_lines(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> list[str]:
    words = text.split()
    lines, current = [], ""
    for word in words:
        test = word if not current else f"{current} {word}"
        if draw.textbbox((0, 0), test, font=font)[2] <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]

def make_intro_page(title: str, logo_path: Path | None, subtitle: str = "") -> Image.Image:
    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(page)
    draw.rectangle((0, 0, PAGE_W, 95), fill=TITLE_COLOR)
    draw.text((120, 210), title, fill=TITLE_COLOR, font=load_font(54, bold=True))
    draw.text((120, 300), subtitle, fill=TEXT_COLOR, font=load_font(28))
    if logo_path and logo_path.exists():
        logo = Image.open(logo_path).convert("RGBA")
        logo.thumbnail((300, 300), Image.LANCZOS)
        page.paste(logo.convert("RGB"), (1450, 150), logo)
    return page

def make_sheet_page(sheet_img_path: Path, sheet_name: str, logo_path: Path | None, show_sheet_name: bool) -> Image.Image:
    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(page)
    sidebar_x = PAGE_W - SIDEBAR_W
    draw.rectangle((sidebar_x, 0, PAGE_W, PAGE_H), fill=SIDEBAR_BG)
    draw.line((sidebar_x, 20, sidebar_x, PAGE_H - 20), fill=DIVIDER, width=2)

    if logo_path and logo_path.exists():
        logo = Image.open(logo_path).convert("RGBA")
        logo.thumbnail((120, 120), Image.LANCZOS)
        x = sidebar_x + (SIDEBAR_W - logo.width) // 2
        page.paste(logo.convert("RGB"), (x, 35), logo)

    if show_sheet_name:
        font = load_font(18, bold=True)
        label = "Nitty Gritty" if sheet_name == "NittyGrittySheet" else sheet_name
        y = 180
        for line in fit_text_lines(draw, label, font, SIDEBAR_W - 16):
            box = draw.textbbox((0, 0), line, font=font)
            draw.text((sidebar_x + (SIDEBAR_W - (box[2] - box[0])) / 2, y), line, fill=TITLE_COLOR, font=font)
            y += 28

    shot = Image.open(sheet_img_path).convert("RGB")
    scale = min(CONTENT_W / shot.width, CONTENT_H / shot.height)
    shot = shot.resize((int(shot.width * scale), int(shot.height * scale)), Image.LANCZOS)
    left = MARGIN_X + (CONTENT_W - shot.width) // 2
    top = MARGIN_Y + (CONTENT_H - shot.height) // 2
    page.paste(shot, (left, top))
    return page
