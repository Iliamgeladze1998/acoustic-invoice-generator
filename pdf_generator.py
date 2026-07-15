"""
PDF Invoice generator: uses acoustic_invoice_final_printable.pdf as a template
and overlays only dynamic content (invoice number, date, buyer, products, total).
All static content stays exactly as in the template.
"""
from __future__ import annotations

import os
import re
import time
import json
import random
import asyncio
from datetime import datetime
from typing import Any
import pytz

import aiohttp
import fitz  # PyMuPDF

from config import OUTPUT_DIR, PRODUCTS_JSON_URL

INVOICE_NUMBERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "invoice_numbers.json")
TEMPLATE_PDF = os.path.join(os.path.dirname(os.path.abspath(__file__)), "acoustic_invoice_final_printable.pdf")

# Cache for the products payload
_PRODUCTS_CACHE: dict[str, Any] | None = None
_PRODUCTS_CACHE_TS: float = 0.0
_PRODUCTS_LOCK = asyncio.Lock()
CACHE_TTL = 600  # 10 minutes


async def _load_products(force: bool = False) -> Any:
    global _PRODUCTS_CACHE, _PRODUCTS_CACHE_TS
    async with _PRODUCTS_LOCK:
        expired = (time.monotonic() - _PRODUCTS_CACHE_TS) > CACHE_TTL
        if _PRODUCTS_CACHE is None or expired or force:
            timeout = aiohttp.ClientTimeout(total=30)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(PRODUCTS_JSON_URL) as resp:
                    resp.raise_for_status()
                    _PRODUCTS_CACHE = await resp.json(content_type=None)
                    _PRODUCTS_CACHE_TS = time.monotonic()
        return _PRODUCTS_CACHE


def _iter_products(payload: Any):
    if isinstance(payload, list):
        for item in payload:
            if isinstance(item, dict):
                yield item
    elif isinstance(payload, dict):
        if "products" in payload and isinstance(payload["products"], list):
            for item in payload["products"]:
                if isinstance(item, dict):
                    yield item
            return
        for key, item in payload.items():
            if isinstance(item, dict):
                item.setdefault("_key", key)
                yield item


_PRODUCT_ID_RE = re.compile(r"product_id=(\d+)")


def _extract_ids(product: dict) -> list[str]:
    ids: list[str] = []
    sku = product.get("sku")
    if sku not in (None, ""):
        ids.append(str(sku).strip())

    url = product.get("url") or product.get("URL") or ""
    if isinstance(url, str):
        m = _PRODUCT_ID_RE.search(url)
        if m:
            ids.append(m.group(1))

    for k in ("product_id", "productId", "id", "ID", "Id", "code", "_key"):
        v = product.get(k)
        if v not in (None, ""):
            ids.append(str(v).strip())

    seen: set[str] = set()
    out: list[str] = []
    for i in ids:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def _extract_name(product: dict) -> str:
    for k in ("product", "title", "name", "full_name", "product_name", "Name", "Title"):
        v = product.get(k)
        if isinstance(v, str) and v.strip():
            return v.strip()
    return ""


def _extract_price(product: dict) -> float:
    for k in ("price", "Price", "unit_price", "sale_price", "current_price", "final_price"):
        v = product.get(k)
        if v in (None, ""):
            continue
        try:
            if isinstance(v, str):
                cleaned = re.sub(r"[^\d.,-]", "", v).replace(",", ".")
                return float(cleaned)
            return float(v)
        except (TypeError, ValueError):
            continue
    return 0.0


def _normalize_code(code: str) -> str:
    return re.sub(r'[^A-Za-z0-9]', '', code).upper()


async def fetch_product_data(product_id: str | int) -> dict:
    target = str(product_id).strip()
    if not target:
        raise ValueError("Product code cannot be empty")

    target_norm = _normalize_code(target)
    payload = await _load_products()

    products_with_ids = []
    for product in _iter_products(payload):
        for pid in _extract_ids(product):
            pid_norm = _normalize_code(pid)
            products_with_ids.append({
                "id": pid,
                "id_norm": pid_norm,
                "name": _extract_name(product),
                "price": _extract_price(product),
            })

    for item in products_with_ids:
        if item["id_norm"] == target_norm:
            return {
                "id": item["id"],
                "name": item["name"],
                "price": item["price"],
            }

    candidates = []
    for item in products_with_ids:
        if item["id_norm"].startswith(target_norm):
            candidates.append(item)

    if not candidates:
        raise ValueError(f"Error: Product code {target} not found in database.")

    if len(candidates) > 1:
        variant_list = ", ".join([f"{c['id']} ({c['name']})" for c in candidates])
        raise ValueError(f"Found multiple variants for this code: {variant_list}. Please specify which one you meant.")

    return {
        "id": candidates[0]["id"],
        "name": candidates[0]["name"],
        "price": candidates[0]["price"],
    }


def _sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|\r\n\t]+", " ", name).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "Client"


def _generate_invoice_number() -> str:
    used = set()
    if os.path.exists(INVOICE_NUMBERS_FILE):
        try:
            with open(INVOICE_NUMBERS_FILE, "r", encoding="utf-8") as f:
                used = set(json.load(f))
        except (json.JSONDecodeError, IOError):
            pass

    for _ in range(10000):
        num = f"INV-{random.randint(100000, 999999)}"
        if num not in used:
            used.add(num)
            with open(INVOICE_NUMBERS_FILE, "w", encoding="utf-8") as f:
                json.dump(sorted(used), f, indent=2)
            return num

    raise RuntimeError("Could not generate a unique invoice number after 10000 attempts.")


# --- PDF Generation using template overlay ---

_FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
_FONT_BOLD_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

# Extract ₾ font from template (Noto-Sans-Bold subset, xref=32)
_GEL_FONT_BUFFER = None

def _get_gel_font_buffer():
    global _GEL_FONT_BUFFER
    if _GEL_FONT_BUFFER is not None:
        return _GEL_FONT_BUFFER
    doc = fitz.open(TEMPLATE_PDF)
    font_data = doc.extract_font(32)
    _GEL_FONT_BUFFER = font_data[3]  # fontbuffer
    doc.close()
    return _GEL_FONT_BUFFER

# Template coordinates (from PyMuPDF text extraction)
# Page size: 595.28 x 841.89 (A4 in points)
POS_ITEM_NUM = 54.2
POS_ITEM_NAME = 80.3
POS_ITEM_QTY = 371.7
POS_ITEM_PRICE_NUM_END = 461.0
POS_ITEM_TOTAL_NUM_END = 536.0
POS_ITEM_PRICE_GEL = 463.1
POS_ITEM_TOTAL_GEL = 538.4
POS_ITEM_ROW_START_Y = 309.1
POS_ITEM_ROW_SPACING = 31.4
POS_GRAND_TOTAL_Y = 691.4
POS_GRAND_TOTAL_GEL = 528.0
POS_GRAND_TOTAL_NUM_END = 525.0

# Table grid line positions (y values for horizontal lines)
# Rows alternate: even rows have fill, odd rows don't
TABLE_LINE_COLOR = (0.886, 0.910, 0.941)
TABLE_ROW_FILL = (0.969, 0.980, 0.988)
TABLE_X_SEGMENTS = [(42.5, 71.3), (71.3, 336.5), (336.5, 412.3), (412.3, 477.5), (477.5, 552.8)]

# Bottom section positions (in template with 11 items)
BOTTOM_LINE_Y = 646.9
BOTTOM_BOX_LEFT = (42.5, 662.3, 308.1, 746.9)
BOTTOM_BOX_RIGHT = (323.1, 662.3, 552.8, 723.9)
BOTTOM_BOX_LEFT_FILL = (0.969, 0.980, 0.988)
BOTTOM_BOX_LEFT_BORDER = (0.886, 0.910, 0.941)
BOTTOM_BOX_RIGHT_FILL = (0.922, 0.973, 0.980)
BOTTOM_BOX_RIGHT_BORDER = (0.0, 0.506, 0.541)

# Bank details text positions (relative to BOTTOM_LINE_Y)
BANK_LABEL_Y = 673.3  # საბანკო რეკვიზიტები:
BANK1_Y = 692.2       # საქართველოს ბანკი:
BANK2_Y = 707.6       # თიბისი ბანკი:
BANK3_Y = 722.9       # კრედო ბანკი:

# Grand total positions (relative to BOTTOM_LINE_Y)
GRAND_TOTAL_LABEL_Y = 672.6
GRAND_TOTAL_NUM_Y = 691.4
GRAND_TOTAL_GEL_Y = 688.8

FOOTER_Y = 816.0
NUM_TEMPLATE_ROWS = 11

# Char width approximations for DejaVuSans at fontsize=10
_CHAR_WIDTH = 5.5
_CHAR_WIDTH_BOLD_16 = 9.0


def _fmt_price_no_currency(val: float) -> str:
    return f"{val:,.0f}".replace(",", " ")


def _split_name(name: str, max_chars: int = 52) -> list[str]:
    """Split a product name into lines that fit within the column."""
    if len(name) <= max_chars:
        return [name]
    # Find a space near the middle
    mid = max_chars
    for i in range(max_chars, max(0, max_chars - 15), -1):
        if i < len(name) and name[i] == ' ':
            mid = i
            break
    return [name[:mid].strip(), name[mid:].strip()]


async def generate_invoice(client_info: str, items: list[dict]) -> str:
    """
    Generate a PDF invoice by overlaying dynamic content on the template.
    Same interface as generator.generate_invoice.

    items: [{'id': '13333', 'qty': 2}, ...]
    Returns absolute path of the generated PDF file.
    """
    if not os.path.exists(TEMPLATE_PDF):
        raise FileNotFoundError(f"Template PDF not found: {TEMPLATE_PDF}")

    # Fetch all products concurrently
    await _load_products()
    resolved = await asyncio.gather(
        *(fetch_product_data(it["id"]) for it in items),
        return_exceptions=True
    )

    for i, result in enumerate(resolved):
        if isinstance(result, Exception):
            raise result

    # Generate invoice number
    invoice_num = _generate_invoice_number()
    invoice_num_short = invoice_num.replace("INV-", "")

    # Date
    tbilisi_tz = pytz.timezone('Asia/Tbilisi')
    current_date = datetime.now(tbilisi_tz).strftime("%d/%m/%Y")

    # Calculate grand total
    grand_total = sum(item.get("qty", 0) * prod["price"] for item, prod in zip(items, resolved))

    # Open template
    doc = fitz.open(TEMPLATE_PDF)
    page = doc[0]

    font_name = "DejaVuSans"
    font_bold = "DejaVuSans-Bold"
    font_gel = "NotoGEL"

    # Register fonts
    page.insert_font(fontname=font_name, fontfile=_FONT_PATH)
    page.insert_font(fontname=font_bold, fontfile=_FONT_BOLD_PATH)
    gel_buffer = _get_gel_font_buffer()
    page.insert_font(fontname=font_gel, fontbuffer=gel_buffer)

    # 1. Cover old invoice number + date
    page.draw_rect(fitz.Rect(350, 95, 570, 120), color=(1,1,1), fill=(1,1,1), overlay=True)
    page.insert_text((363.2, 105.0), f"# {invoice_num}  |  თარიღი: {current_date}",
                     fontname=font_name, fontsize=10, color=(0,0,0))

    # 2. Cover old buyer name area
    page.draw_rect(fitz.Rect(310, 160, 570, 205), color=(1,1,1), fill=(1,1,1), overlay=True)
    # Write buyer name — split into lines if too long
    buyer_lines = _split_name(client_info, max_chars=48)
    by = 170.3
    for bl in buyer_lines:
        page.insert_text((313.4, by), bl,
                         fontname=font_bold, fontsize=10, color=(0,0,0))
        by += 14

    # 3. Cover everything from below table header to footer
    page.draw_rect(fitz.Rect(35, 290, 560, 760), color=(1,1,1), fill=(1,1,1), overlay=True)

    # Calculate shift: how much to move the bottom section up
    num_items = len(items)
    shift = (NUM_TEMPLATE_ROWS - num_items) * POS_ITEM_ROW_SPACING

    # 4. Draw table grid lines and alternating row backgrounds for N rows
    for row_i in range(num_items):
        row_y = 295 + row_i * POS_ITEM_ROW_SPACING
        row_y_next = 295 + (row_i + 1) * POS_ITEM_ROW_SPACING

        # Alternating row fill (even rows get fill in template: rows 0,2,4,6,8)
        if row_i % 2 == 0:
            for x0, x1 in TABLE_X_SEGMENTS:
                page.draw_rect(fitz.Rect(x0, row_y, x1, row_y_next),
                               color=None, fill=TABLE_ROW_FILL, overlay=True)

        # Horizontal grid line at bottom of row
        for x0, x1 in TABLE_X_SEGMENTS:
            page.draw_line(fitz.Point(x0, row_y_next), fitz.Point(x1, row_y_next),
                           color=TABLE_LINE_COLOR, width=0.5, overlay=True)

    # 5. Draw bottom section at shifted position
    bottom_line_y = BOTTOM_LINE_Y - shift

    # Horizontal line above bottom section
    for x0, x1 in TABLE_X_SEGMENTS:
        page.draw_line(fitz.Point(x0, bottom_line_y), fitz.Point(x1, bottom_line_y),
                       color=TABLE_LINE_COLOR, width=0.5, overlay=True)

    # Left box (bank details)
    bl = BOTTOM_BOX_LEFT
    bl_new = (bl[0], bl[1] - shift, bl[2], bl[3] - shift)
    page.draw_rect(fitz.Rect(*bl_new), color=None, fill=BOTTOM_BOX_LEFT_FILL, overlay=True)
    page.draw_rect(fitz.Rect(*bl_new), color=BOTTOM_BOX_LEFT_BORDER, width=1, fill=None, overlay=True)

    # Right box (grand total)
    br = BOTTOM_BOX_RIGHT
    br_new = (br[0], br[1] - shift, br[2], br[3] - shift)
    page.draw_rect(fitz.Rect(*br_new), color=None, fill=BOTTOM_BOX_RIGHT_FILL, overlay=True)
    page.draw_rect(fitz.Rect(*br_new), color=BOTTOM_BOX_RIGHT_BORDER, width=1, fill=None, overlay=True)

    # Bank details text
    bank_y = BANK_LABEL_Y - shift
    page.insert_text((52.3, bank_y), "საბანკო რეკვიზიტები:",
                     fontname=font_bold, fontsize=9, color=(0,0,0))
    bank1_y = BANK1_Y - shift
    page.insert_text((52.3, bank1_y), "საქართველოს ბანკი:",
                     fontname=font_bold, fontsize=8, color=(0,0,0))
    page.insert_text((140.4, bank1_y), "GE75BG0000000346094490",
                     fontname=font_name, fontsize=8, color=(0,0,0))
    bank2_y = BANK2_Y - shift
    page.insert_text((52.3, bank2_y), "თიბისი ბანკი:",
                     fontname=font_bold, fontsize=8, color=(0,0,0))
    page.insert_text((109.9, bank2_y), "GE89TB7474136080100007",
                     fontname=font_name, fontsize=8, color=(0,0,0))
    bank3_y = BANK3_Y - shift
    page.insert_text((52.3, bank3_y), "კრედო ბანკი:",
                     fontname=font_bold, fontsize=8, color=(0,0,0))
    page.insert_text((109.4, bank3_y), "GE33CD0360001020383062",
                     fontname=font_name, fontsize=8, color=(0,0,0))

    # Grand total label + amount + ₾
    gt_label_y = GRAND_TOTAL_LABEL_Y - shift
    page.insert_text((360.2, gt_label_y), "სულ გადასახდელი (დღგ-ს ჩათვლით):",
                     fontname=font_bold, fontsize=10, color=(0,0,0))
    gt_num_y = GRAND_TOTAL_NUM_Y - shift
    gt_gel_y = GRAND_TOTAL_GEL_Y - shift
    grand_total_str = _fmt_price_no_currency(grand_total)
    gt_width = len(grand_total_str) * _CHAR_WIDTH_BOLD_16
    page.insert_text((POS_GRAND_TOTAL_NUM_END - gt_width, gt_num_y), grand_total_str,
                     fontname=font_bold, fontsize=16, color=(0,0,0))
    page.insert_text((POS_GRAND_TOTAL_GEL, gt_gel_y), "₾",
                     fontname=font_gel, fontsize=16, color=(0,0,0))

    # 6. Write product rows
    y = POS_ITEM_ROW_START_Y
    for idx, (item, prod) in enumerate(zip(items, resolved), start=1):
        qty = item.get("qty", 0)
        try:
            qty = int(qty)
        except (TypeError, ValueError):
            qty = 0

        price = prod["price"]
        line_total = qty * price

        # №
        page.insert_text((POS_ITEM_NUM, y), str(idx),
                         fontname=font_name, fontsize=10, color=(0,0,0))

        # დასახელება
        name_lines = _split_name(prod["name"], max_chars=52)
        if len(name_lines) == 1:
            page.insert_text((POS_ITEM_NAME, y), name_lines[0],
                             fontname=font_name, fontsize=10, color=(0,0,0))
        else:
            page.insert_text((POS_ITEM_NAME, y - 3), name_lines[0],
                             fontname=font_name, fontsize=10, color=(0,0,0))
            page.insert_text((POS_ITEM_NAME, y + 10), name_lines[1],
                             fontname=font_name, fontsize=10, color=(0,0,0))

        # რაოდენობა
        page.insert_text((POS_ITEM_QTY, y), str(qty),
                         fontname=font_name, fontsize=10, color=(0,0,0))

        # ფასი
        price_str = _fmt_price_no_currency(price)
        text_width = len(price_str) * _CHAR_WIDTH
        page.insert_text((POS_ITEM_PRICE_NUM_END - text_width, y), price_str,
                         fontname=font_name, fontsize=10, color=(0,0,0))
        page.insert_text((POS_ITEM_PRICE_GEL, y), "₾",
                         fontname=font_gel, fontsize=10, color=(0,0,0))

        # ჯამი
        total_str = _fmt_price_no_currency(line_total)
        text_width2 = len(total_str) * _CHAR_WIDTH
        page.insert_text((POS_ITEM_TOTAL_NUM_END - text_width2, y), total_str,
                         fontname=font_name, fontsize=10, color=(0,0,0))
        page.insert_text((POS_ITEM_TOTAL_GEL, y), "₾",
                         fontname=font_gel, fontsize=10, color=(0,0,0))

        y += POS_ITEM_ROW_SPACING

    # 7. Merge page 2 content onto page 1 if there's enough space
    PAGE2_CONTENT_TOP = 70.0
    PAGE2_CONTENT_BOT = 270.0
    PAGE2_CONTENT_HEIGHT = PAGE2_CONTENT_BOT - PAGE2_CONTENT_TOP

    bottom_end_y = BOTTOM_BOX_LEFT[3] - shift  # 746.9 - shift
    available = 841.89 - bottom_end_y - 15  # margin

    if doc.page_count > 1 and available >= PAGE2_CONTENT_HEIGHT:
        page2 = doc[1]
        # Place page 2 content right after bottom section
        target_y = bottom_end_y + 15
        offset_y = target_y - PAGE2_CONTENT_TOP

        # Copy text spans from page 2 to page 1
        blocks2 = page2.get_text('dict')['blocks']
        for b in blocks2:
            if 'lines' not in b:
                continue
            for line in b['lines']:
                for span in line['spans']:
                    t = span['text']
                    if not t.strip() or 'გვერდი' in t:
                        continue
                    sx = span['bbox'][0]
                    sy = span['bbox'][1]
                    sz = span['size']
                    sf = span['font']
                    # Map font names
                    fn = font_name
                    if 'Bold' in sf:
                        fn = font_bold
                    elif 'Noto' in sf:
                        fn = font_gel
                    page.insert_text((sx, sy + offset_y), t,
                                     fontname=fn, fontsize=sz, color=(0,0,0))

        # Copy images from page 2 to page 1
        for img in page2.get_images():
            xref = img[0]
            rects = page2.get_image_rects(xref)
            for r in rects:
                # Extract image and insert at shifted position
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                img_rect = fitz.Rect(r.x0, r.y0 + offset_y, r.x1, r.y1 + offset_y)
                page.insert_image(img_rect, pixmap=pix)
                pix = None

        # Delete page 2
        doc.delete_page(1)

    # Save
    safe_name = _sanitize_filename(client_info)
    out_path = os.path.abspath(os.path.join(OUTPUT_DIR, f"Invoice_{invoice_num}_{safe_name}.pdf"))
    doc.save(out_path)
    doc.close()
    return out_path
