import os

TELEGRAM_TOKEN = "8618499132:AAGKX1JDXp-gopw6d3CYeEYCfcFRjB2DQbI"
TEMPLATE_PATH = "/root/Acoustic-Invoice-Bot/ინვოისი - Marneuli.xlsx"
OUTPUT_DIR = "/root/Acoustic-Invoice-Bot/generated_invoices"
PRODUCTS_JSON_URL = "https://acoustic.ge/data/products.json"

os.makedirs(OUTPUT_DIR, exist_ok=True)
