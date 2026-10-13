"""Local smoke test: pull demo data via SQL API into pandas, build decks, sanity-check them."""
import json, subprocess, sys
import pandas as pd
sys.path.insert(0, "src")
import build_deck as bd

def q(sql):
    body = {"warehouse_id": "7061d004ec1f1f1e", "wait_timeout": "50s", "statement": sql}
    r = subprocess.run(["databricks", "api", "post", "/api/2.0/sql/statements", "--json", json.dumps(body)],
                       capture_output=True, text=True, check=True)
    d = json.loads(r.stdout)
    assert d["status"]["state"] == "SUCCEEDED", d["status"]
    cols = [c["name"] for c in d["manifest"]["schema"]["columns"]]
    return pd.DataFrame(d["result"]["data_array"], columns=cols)

C = "fd_fevm_classic_catalog.aibi"
o = q(f"""SELECT o.order_id, o.order_date, c.customer_id, c.customer_name, c.segment, c.region, p.product_name,
 p.category, o.quantity, o.unit_price, o.discount, o.status, o.quantity*o.unit_price*(1-o.discount) revenue
 FROM {C}.fact_orders o JOIN {C}.dim_customers c USING(customer_id) JOIN {C}.dim_products p USING(product_id)""")
for col in ["order_id","customer_id","quantity"]: o[col] = o[col].astype(int)
for col in ["unit_price","discount","revenue"]: o[col] = o[col].astype(float)
inv = q(f"SELECT p.product_name, i.warehouse, i.quantity_on_hand, i.quantity_reserved, i.reorder_point FROM {C}.fact_inventory i JOIN {C}.dim_products p USING(product_id)")
for col in ["quantity_on_hand","quantity_reserved","reorder_point"]: inv[col] = inv[col].astype(int)

for name, sel in {"multi": o, "europe": o[o.region == "Europe"], "na_nopending": o[(o.region == "North America") & (o.status != "Pending")]}.items():
    data, realm, n = bd.build_deck(sel, inv, "2026-01 .. 2026-09")
    open(f"/tmp/deck_{name}.pptx", "wb").write(data)
    print(name, realm, "slides", n, "bytes", len(data), "orders", sel.order_id.nunique())
