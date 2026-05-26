# Inventory Reconciler

A local web tool for reconciling physical inventory against an on-hand sheet and Fixably export.

## What it does

- **Step 1** — Compare physically scanned serials vs your Actual / On Hand sheet
- **Step 2** — Compare your on-hand sheet vs Fixably export to find discrepancies
- **Stocked Out** — Upload a stocked-out sheet to exclude serials already used but not yet updated in Fixably
- **Export** — Generate a CSV report with actionable items (Add to Fixably, Remove from Fixably, Update qty)

## Stack

- Python / Flask
- Pandas
- Vanilla JS + HTML

## Run

```bash
pip install -r requirements.txt
python app.py
```

Open http://localhost:5050
