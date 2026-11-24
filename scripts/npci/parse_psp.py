"""Build data/npci_payer_psp_declines_monthly.csv from the raw NPCI "Top 15 PSPs" (payer) responses.

Usage: python parse_psp.py ../../data/npci_top15_payer_psps_raw.json ../../data/npci_payer_psp_declines_monthly.csv
"""
import csv, json, re, sys

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
# NPCI spells the same bank several ways ("Yes Bank Ltd.", "YES BANK LTD", ...); map the big ones to one name.
CANON = [("yes bank", "Yes Bank"), ("axis bank", "Axis Bank"), ("hdfc bank", "HDFC Bank"), ("icici bank", "ICICI Bank"),
         ("state bank of india", "SBI"), ("kotak mahindra", "Kotak Mahindra Bank")]


def bank(name):
    n = re.sub(r"\s+", " ", name.replace("\xa0", " ")).strip()
    for key, canon in CANON:
        if n.lower().startswith(key):
            return canon
    return n


def num(x):
    return float(str(x).replace(",", "").replace("%", "").strip())


raw, out = sys.argv[1], sys.argv[2]
data = json.load(open(raw))
rows = []
for key, resp in data.items():
    if resp.get("s") != 200:
        continue
    y, m = key.split("-")
    for r in json.loads(resp["t"]).get("data", {}).get("results") or []:
        rows.append([f"{y}-{MONTHS.index(m) + 1:02d}", bank(r["payer_psp"]), num(r["total_volume_in_mn"]),
                     num(r["approved_percent"]), num(r["bd_percent"]), num(r["td_percent"])])
rows.sort()
with open(out, "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["month", "psp", "vol_mn", "approved", "bd", "td"])
    w.writerows(rows)
print(len(rows), "rows,", len({r[0] for r in rows}), "months")
