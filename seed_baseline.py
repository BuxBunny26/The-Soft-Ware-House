"""Seed March 2026 as the first monthly snapshot for comparison baseline."""
import os, json, shutil

BASE = os.path.dirname(os.path.abspath(__file__))
MONTHLY_DIR = os.path.join(BASE, "Monthly_Data")
MARCH_DIR = os.path.join(MONTHLY_DIR, "2026-03")
os.makedirs(MARCH_DIR, exist_ok=True)

# Copy the user CSV
src_csv = os.path.join(BASE, "Source Data", "users_2026_03_30 11_05_42.csv")
if os.path.exists(src_csv):
    shutil.copy2(src_csv, os.path.join(MARCH_DIR, "users_2026_03_30 11_05_42.csv"))
    print("Copied user CSV")

# Copy the invoice PDF if it exists
for f in os.listdir(os.path.join(BASE, "Source Data")):
    if f.lower().endswith(".pdf") and "inv" in f.lower():
        shutil.copy2(os.path.join(BASE, "Source Data", f), os.path.join(MARCH_DIR, f))
        print(f"Copied invoice: {f}")

# Write metadata
meta = {
    "month": "2026-03 (March 2026)",
    "uploaded": "2026-03-30 (baseline)",
    "files": os.listdir(MARCH_DIR),
}
with open(os.path.join(MARCH_DIR, "meta.json"), "w") as f:
    json.dump(meta, f, indent=2)
print("Wrote meta.json")

# Build summary from the CSV
import sys
sys.path.insert(0, BASE)
from dashboard import build_month_summary
summary = build_month_summary(MARCH_DIR)
print(f"Built summary: {summary['total_users']} users, {len(summary['sku_counts'])} SKUs")
print("Done - March 2026 baseline seeded")
