"""Quick verification of dashboard numbers after cost model change."""
import urllib.request
import re

html = urllib.request.urlopen("http://localhost:5000/").read().decode()

# Find key values
patterns = [
    ("Invoice Total", r'Invoice Total.*?kpi-value.*?>(R[\d,\.]+)<'),
    ("Allocated to Users", r'Allocated to Users.*?kpi-value.*?>(R[\d,\.]+)<'),
    ("Unallocated/Wasted", r'Unallocated.*?kpi-value.*?>(R[\d,\.]+)<'),
    ("Net Unallocated", r'NET UNALLOCATED.*?text-danger.*?>(R[\d,\.]+)<'),
]
for label, pat in patterns:
    m = re.search(pat, html, re.DOTALL)
    print(f"{label}: {m.group(1) if m else 'NOT FOUND'}")

# Waste items (red rows = wasted)
print("\nWaste breakdown (from overview):")
red_rows = re.findall(r'table-danger.*?<td class="ps-3"><small>(.*?)</small></td>.*?<span class="text-danger">(R[\d,\.]+)</span>', html, re.DOTALL)
for sku, amt in red_rows:
    print(f"  WASTED: {sku} = {amt}")

green_rows = re.findall(r'table-success.*?<td class="ps-3"><small>(.*?)</small></td>.*?<span class="text-success">-(R[\d,\.]+)</span>', html, re.DOTALL)
for sku, amt in green_rows:
    print(f"  SAVING: {sku} = -{amt}")

# Check users page for sample costs
html2 = urllib.request.urlopen("http://localhost:5000/users").read().decode()
# Count cost values
costs = re.findall(r'class="text-end">(R[\d,\.]+)</td>', html2)
print(f"\nUsers page: {len(costs)} cost values found")
if costs:
    print(f"  First 5: {costs[:5]}")
    print(f"  Last 5: {costs[-5:]}")
