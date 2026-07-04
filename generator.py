"""
Invoice generator: fetches live product data from acoustic.ge JSON feed
and populates the master Excel template using a clean-write-save approach.
"""
from __future__ import annotations

import os
import re
import time
import asyncio
from datetime import datetime
from typing import Any
import pytz

import aiohttp
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.cell.cell import MergedCell

from config import TEMPLATE_PATH, OUTPUT_DIR, PRODUCTS_JSON_URL


def _safe_cell(ws, row, col):
    """Return the writable cell at (row, col), routing MergedCell to the range's top-left."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return cell
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell


def _write_cell(ws, row, col, value=None, number_format=None):
    """Write value / number_format to a cell, transparently handling merged ranges."""
    cell = _safe_cell(ws, row, col)
    if value is not None:
        cell.value = value
    if number_format is not None:
        cell.number_format = number_format

# Cache for the products payload so repeated lookups inside one
# generate_invoice() call don't hammer the API.
# Expires after CACHE_TTL seconds so long-running bot always serves fresh prices.
_PRODUCTS_CACHE: dict[str, Any] | None = None
_PRODUCTS_CACHE_TS: float = 0.0
_PRODUCTS_LOCK = asyncio.Lock()
CACHE_TTL = 600  # 10 minutes


async def _load_products(force: bool = False) -> Any:
    """Fetch the full products JSON and cache it in memory (with TTL)."""
    global _PRODUCTS_CACHE, _PRODUCTS_CACHE_TS
    async with _PRODUCTS_LOCK:
        expired = (time.monotonic() - _PRODUCTS_CACHE_TS) > CACHE_TTL
        if _PRODUCTS_CACHE is None or expired or force:
            timeout = aiohttp.ClientTimeout(total=30)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(PRODUCTS_JSON_URL) as resp:
                    resp.raise_for_status()
                    # content_type may not be application/json -> force parse
                    _PRODUCTS_CACHE = await resp.json(content_type=None)
                    _PRODUCTS_CACHE_TS = time.monotonic()
        return _PRODUCTS_CACHE


def _iter_products(payload: Any):
    """Yield individual product dicts from whatever shape the JSON has."""
    if isinstance(payload, list):
        for item in payload:
            if isinstance(item, dict):
                yield item
    elif isinstance(payload, dict):
        # Common shapes: {"products": [...]} or {"<id>": {...}, ...}
        if "products" in payload and isinstance(payload["products"], list):
            for item in payload["products"]:
                if isinstance(item, dict):
                    yield item
            return
        for key, item in payload.items():
            if isinstance(item, dict):
                # Inject the key as a candidate id if missing
                item.setdefault("_key", key)
                yield item


_PRODUCT_ID_RE = re.compile(r"product_id=(\d+)")


def _extract_ids(product: dict) -> list[str]:
    """
    Return ALL identifiers a product can be matched on.
    The acoustic.ge feed exposes two distinct numeric IDs:
      - `sku`          (top-level, what staff usually quotes)
      - `product_id`   (extracted from the URL)
    Plus fallbacks for other JSON shapes.
    """
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

    # de-dup while preserving order
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
    """Remove all non-alphanumeric characters and normalize case for flexible matching."""
    return re.sub(r'[^A-Za-z0-9]', '', code).upper()


async def fetch_product_data(product_id: str | int) -> dict:
    """
    Return {'id', 'name', 'price'} for the requested product_id.
    Raises ValueError if not found or if multiple variants exist.
    Uses character-agnostic normalization for flexible matching.
    """
    target = str(product_id).strip()
    if not target:
        raise ValueError("Product code cannot be empty")

    # Normalize target (remove all non-alphanumeric characters)
    target_norm = _normalize_code(target)

    payload = await _load_products()

    # Build a list of all products with their normalized IDs
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

    # Pass B: Look for exact match after normalization
    for item in products_with_ids:
        if item["id_norm"] == target_norm:
            return {
                "id": item["id"],
                "name": item["name"],
                "price": item["price"],
            }

    # Pass C: Look for startswith (normalized input)
    candidates = []
    for item in products_with_ids:
        if item["id_norm"].startswith(target_norm):
            candidates.append(item)

    if not candidates:
        raise ValueError(f"Error: Product code {target} not found in database.")

    # Pass D: If multiple matches, ask user to clarify
    if len(candidates) > 1:
        variant_list = ", ".join([f"{c['id']} ({c['name']})" for c in candidates])
        raise ValueError(f"Found multiple variants for this code: {variant_list}. Please specify which one you meant.")

    # Pass E: Only one match, use it automatically
    return {
        "id": candidates[0]["id"],
        "name": candidates[0]["name"],
        "price": candidates[0]["price"],
    }


def _sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|\r\n\t]+", " ", name).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "Client"


async def generate_invoice(client_info: str, items: list[dict]) -> str:
    """
    Populate the master template with the given client + items, save the
    result into OUTPUT_DIR, and return the absolute path of the new file.

    items: [{'id': '13333', 'qty': 2}, ...]
    """
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")

    # The template's item table holds a fixed number of rows; writing more
    # would overwrite the Grand Total / footer cells below it.
    MAX_ITEMS = 4
    if len(items) > MAX_ITEMS:
        raise ValueError(
            f"შაბლონში მაქსიმუმ {MAX_ITEMS} პროდუქტი ეტევა, შენ გამოგზავნე {len(items)}. "
            f"გთხოვ გაყო ინვოისი რამდენიმე ნაწილად."
        )

    # Fetch all products concurrently (single HTTP call, just lookups in cache)
    await _load_products()
    resolved = await asyncio.gather(
        *(fetch_product_data(it["id"]) for it in items),
        return_exceptions=True
    )

    # Check for errors (ValueError from fetch_product_data)
    for i, result in enumerate(resolved):
        if isinstance(result, Exception):
            raise result  # Re-raise the ValueError with the specific message

    # Load template directly (no shutil.copy)
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Client name -> A3 (transparently handles merged cells)
    _write_cell(ws, 3, 1, value=client_info)

    # Items table -> starting row 8, clear rows 8-11 before writing
    START_ROW = 8
    ITEM_TABLE_ROWS = 4  # rows 8, 9, 10, 11

    # Clear item table cells (A, B, I, J, K) by setting to empty string
    for r in range(START_ROW, START_ROW + ITEM_TABLE_ROWS):
        for c in (1, 2, 9, 10, 11):
            _write_cell(ws, r, c, value="")

    # Write product data
    for idx, (item, prod) in enumerate(zip(items, resolved), start=1):
        row = START_ROW + idx - 1
        qty = item.get("qty", 0)
        try:
            qty = int(qty)
        except (TypeError, ValueError):
            qty = 0

        price = prod["price"]
        line_total = qty * price

        _write_cell(ws, row, 1, value=idx)              # A: counter
        _write_cell(ws, row, 2, value=prod["name"])     # B: full title
        _write_cell(ws, row, 9, value=qty)              # I: qty
        _write_cell(ws, row, 10, value=price, number_format='#,##0 ₾')   # J: unit price
        _write_cell(ws, row, 11, value=line_total, number_format='#,##0 ₾')  # K: line total

    # Auto-fit column widths for J and K
    ws.column_dimensions['J'].width = 15
    ws.column_dimensions['K'].width = 15

    # Calculate and write Grand Total to K13
    grand_total = sum(item.get("qty", 0) * prod["price"] for item, prod in zip(items, resolved))
    _write_cell(ws, 13, 11, value=grand_total, number_format='#,##0 ₾')  # K13: Grand Total

    # Update date in K18 with Asia/Tbilisi timezone
    tbilisi_tz = pytz.timezone('Asia/Tbilisi')
    current_date = datetime.now(tbilisi_tz).strftime("%d/%m/%y")
    _write_cell(ws, 18, 11, value=f"თარიღი: {current_date}")

    # Load images from project folder with exact positioning
    # Remove any existing images to avoid duplicates
    if hasattr(ws, "_images"):
        del ws._images[:]
    
    # Load template to get original image anchors
    template_wb = openpyxl.load_workbook(TEMPLATE_PATH)
    template_ws = template_wb.active
    
    # Load images from project folder (logo.png and signature.png)
    # with exact positioning matching the template
    
    # Logo: width 689, height 150, col 5, row 0, offset 59194, 301361
    logo_path = os.path.join(os.path.dirname(TEMPLATE_PATH), "logo.png")
    if os.path.exists(logo_path) and len(template_ws._images) >= 2:
        logo_img = Image(logo_path)
        logo_img.width = 689
        logo_img.height = 150
        # Copy anchor from second image in template (logo)
        logo_img.anchor = template_ws._images[1].anchor
        ws.add_image(logo_img)
    
    # Signature: width 139, height 46, col 7, row 16, offset 52351, 177086
    signature_path = os.path.join(os.path.dirname(TEMPLATE_PATH), "signature.png")
    if os.path.exists(signature_path) and len(template_ws._images) >= 1:
        signature_img = Image(signature_path)
        signature_img.width = 139
        signature_img.height = 46
        # Copy anchor from first image in template (signature)
        signature_img.anchor = template_ws._images[0].anchor
        ws.add_image(signature_img)

    # Save as new file
    safe_name = _sanitize_filename(client_info)
    out_path = os.path.abspath(os.path.join(OUTPUT_DIR, f"Invoice_{safe_name}.xlsx"))
    wb.save(out_path)
    return out_path
