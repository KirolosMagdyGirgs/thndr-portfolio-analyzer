# config.py

BASE_URL = "https://web.thndr.app/invest"
CSS_SELECTOR = "[class^='sc-olbas iBxzbw'}"
REQUIRED_KEYS = [
    "Asset Class",   # ← was missing comma before, causing it to merge with next key
    "Units Owned",
    "Cost Per Unit",
    "Current Price",
    "Market Value",
    "Daily Change",
    "Unrealized Return",
]