# Google Pay product teardown: how GPay lost India's UPI lead to PhonePe

**Read the deck:** [gpay-teardown.pdf](gpay-teardown.pdf) (12 slides, November 2026). Editable source: [gpay-teardown.pptx](gpay-teardown.pptx).

## The question
Why did Google Pay lose India's UPI lead to PhonePe, and how does it win it back?

## The answer, in short
- **What happened:** GPay's UPI share fell from **44.3% (Jun 2020) to 32.1% (Oct 2026)**. PhonePe overtook it in **Dec 2020**, and leads with **45.3%**. GPay's volume still grew 13.7x, against 18.9x for the market.
- **Why:** PhonePe won the shop-counter payment habit while GPay pulled back on rewards (2020–21). Tier-2/3 reach and reliability added to it.
- **North-star metric:** **weekly merchant payers**, meaning GPay users who pay a shop in at least 3 of the last 4 weeks. Not share (NPCI caps it at 30%), not volume (cashback splitting inflates it).
- **Options:** 5 options scored two ways, with RICE and a weighted strategic score (moat, revenue, regulation).
- **Plan:** launch habit rewards this quarter, and bring credit to the pay button within a year.

## Method
- 77 months of NPCI app-wise UPI statistics (Jun 2020 – Oct 2026), plus NPCI payer-bank decline data, PhonePe's DRHP, press reports and user and merchant quotes. Every slide lists its sources.
- This is an outside-in teardown from public data only. There is no internal Google data. Segment sizes use PhonePe's published split as a proxy.

## Files
| Path | What it is |
|---|---|
| `gpay-teardown.pdf` / `.pptx` | The deck |
| `data/npci_upi_apps_monthly.csv` | Monthly UPI volume and share by app (from NPCI's public statistics) |
| `data/npci_payer_psp_declines_monthly.csv` | Monthly decline rates of the top payer apps |
| `data/*_raw.json` | Raw NPCI responses the CSVs were built from |
| `scripts/` | Playwright scripts used to collect app-store pages and reviews |
| `deck/build.js` | Builds the deck (pptxgenjs) |

App-store screenshots used in the deck belong to their owners and are not included here.

Author: Mohammad Faiz Ahmad
