import os
import sys

# Ensure scraper_common is importable for products_cache
_scraper_common = "/root/scraper_common"
if _scraper_common not in sys.path:
    sys.path.insert(0, _scraper_common)


def _load_env(path: str) -> None:
    """Minimal .env loader (no external dependency)."""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip())


_load_env(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
if not TELEGRAM_TOKEN:
    raise RuntimeError("TELEGRAM_TOKEN is not set. Add it to .env or the environment.")
TEMPLATE_PATH = "/root/Acoustic-Invoice-Bot/ინვოისი - ნიმუში საბოლოო.xlsx"
OUTPUT_DIR = "/root/Acoustic-Invoice-Bot/generated_invoices"
PRODUCTS_JSON_URL = "https://acoustic.ge/data/products.json"

os.makedirs(OUTPUT_DIR, exist_ok=True)
