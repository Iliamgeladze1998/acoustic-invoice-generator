"""
Invoice generator: fetches live product data from acoustic.ge JSON feed
and populates the master Excel template using a clean-write-save approach.
"""
from __future__ import annotations

import os
import re
import asyncio
from datetime import datetime
from typing import Any

import aiohttp
import openpyxl

from config import TEMPLATE_PATH, OUTPUT_DIR, PRODUCTS_JSON_URL

# Cache for the products payload so repeated lookups inside one
# generate_invoice() call don't hammer the API.
_PRODUCTS_CACHE: dict[str, Any] | None = None
_PRODUCTS_LOCK = asyncio.Lock()


async def _load_products(force: bool = False) -> Any:
    """Fetch the full products JSON once and cache it in memory."""
    global _PRODUCTS_CACHE
    async with _PRODUCTS_LOCK:
        if _PRODUCTS_CACHE is None or force:
            timeout = aiohttp.ClientTimeout(total=30)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(PRODUCTS_JSON_URL) as resp:
                    resp.raise_for_status()
                    # content_type may not be application/json -> force parse
                    _PRODUCTS_CACHE = await resp.json(content_type=None)
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


async def fetch_product_data(product_id: str | int) -> dict | None:
    """
    Return {'id', 'name', 'price'} for the requested product_id, or None
    if not found. Matches robustly across string/int IDs.
    """
    target = str(product_id).strip()
    if not target:
        return None

    payload = await _load_products()
    target_norm = target.lstrip("0") or "0"

    for product in _iter_products(payload):
        for pid in _extract_ids(product):
            if pid == target or pid.lstrip("0") == target_norm:
                return {
                    "id": pid,
                    "name": _extract_name(product),
                    "price": _extract_price(product),
                }
    return None


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

    # Fetch all products concurrently (single HTTP call, just lookups in cache)
    await _load_products()
    resolved = await asyncio.gather(
        *(fetch_product_data(it["id"]) for it in items)
    )

    missing = [items[i]["id"] for i, r in enumerate(resolved) if r is None]
    if missing:
        raise ValueError(
            "The following product IDs were not found in the live feed: "
            + ", ".join(missing)
        )

    # Load template directly (no shutil.copy)
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Client name -> A3
    ws.cell(row=3, column=1, value=client_info)

    # Items table -> starting row 8, clear rows 8-11 before writing
    START_ROW = 8
    ITEM_TABLE_ROWS = 4  # rows 8, 9, 10, 11

    # Clear item table cells (A, B, I, J, K) by setting to empty string
    for r in range(START_ROW, START_ROW + ITEM_TABLE_ROWS):
        for c in (1, 2, 9, 10, 11):
            ws.cell(row=r, column=c, value="")

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

        ws.cell(row=row, column=1, value=idx)              # A: counter
        ws.cell(row=row, column=2, value=prod["name"])     # B: full title
        ws.cell(row=row, column=9, value=qty)              # I: qty
        ws.cell(row=row, column=10, value=price)           # J: unit price
        ws.cell(row=row, column=10).number_format = '#,##0'  # J: numeric format
        ws.cell(row=row, column=11, value=line_total)       # K: line total (calculated in Python)
        ws.cell(row=row, column=11).number_format = '#,##0'  # K: numeric format

    # Auto-fit column widths for J and K
    ws.column_dimensions['J'].width = 15
    ws.column_dimensions['K'].width = 15

    # Calculate and write Grand Total to K13
    grand_total = sum(item.get("qty", 0) * prod["price"] for item, prod in zip(items, resolved))
    ws.cell(row=13, column=11, value=grand_total)  # K13: Grand Total
    ws.cell(row=13, column=11).number_format = '#,##0'  # K13: numeric format

    # Update date in K18
    current_date = datetime.now().strftime("%d/%m/%y")
    ws.cell(row=18, column=11, value=f"თარიღი: {current_date}")

    # Save as new file
    safe_name = _sanitize_filename(client_info)
    out_path = os.path.abspath(os.path.join(OUTPUT_DIR, f"Invoice_{safe_name}.xlsx"))
    wb.save(out_path)
    return out_path
