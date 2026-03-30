"""Re-seed March 2026 baseline with updated parser."""
import importlib, sys, os
sys.path.insert(0, os.getcwd())
import dashboard
importlib.reload(dashboard)

folder = os.path.join(os.getcwd(), "Monthly_Data", "2026-03")
summary = dashboard.build_month_summary(folder)
print(f"Total users: {summary['total_users']}")
print(f"Licensed users: {summary['total_licensed']}")
print(f"SKUs found: {len(summary['sku_counts'])}")
for sku, count in sorted(summary["sku_counts"].items(), key=lambda x: -x[1]):
    print(f"  {sku}: {count}")
