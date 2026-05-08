
<div align="center">

```
██████╗ ██╗██████╗       ██╗  ██╗
██╔══██╗██║██╔══██╗      ╚██╗██╔╝
██████╔╝██║██║  ██║       ╚███╔╝ 
██╔══██╗██║██║  ██║       ██╔██╗ 
██████╔╝██║██████╔╝      ██╔╝ ██╗
╚═════╝ ╚═╝╚═════╝       ╚═╝  ╚═╝
```

### Procurement Intelligence Platform

*Upload supplier bid data · Analyse costs · Rank vendors · Export insights*

---

![Python](https://img.shields.io/badge/Python-3.11+-white?style=flat-square&logo=python&logoColor=white&labelColor=FF8C42&color=FFD8C2)
![Flask](https://img.shields.io/badge/Flask-3.1-white?style=flat-square&logo=flask&logoColor=white&labelColor=FF6B35&color=FFE8DC)
![Pandas](https://img.shields.io/badge/Pandas-2.2-white?style=flat-square&logo=pandas&logoColor=white&labelColor=38BDF8&color=E0F4FF)
![Upload](https://img.shields.io/badge/Input-CSV_Upload-white?style=flat-square&labelColor=0EA5E9&color=E0F4FF)
![Modules](https://img.shields.io/badge/7_Analysis_Modules-white?style=flat-square&labelColor=FB923C&color=FFF0E8)
![License](https://img.shields.io/badge/License-MIT-white?style=flat-square&labelColor=38BDF8&color=E0F7FF)

<br/>

[![Deploy on Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy)

<br/>

</div>

---

## What is BID-X?

**BID-X** is a web-based procurement intelligence tool that lets sourcing and procurement teams upload supplier bid CSV files and instantly run structured analysis — no spreadsheets, no manual pivoting.

> Upload a bid file → preview the data → click any analysis module → get instant results.

---

## Features

| Module | What it does |
|---|---|
| 📊 **Bid Overview** | KPI cards (total bids, lowest/highest, avg unit price) + full bid register |
| 💡 **Unit Price Analysis** | Colour heatmap — green = cheapest, red = most expensive, per product × supplier |
| 💰 **Total Cost Analysis** | Animated bar charts, delivery comparison, cost variance vs market average |
| 📦 **Quantity Analysis** | Compliance check per product/supplier with deviation % and status flags |
| 🏆 **Supplier Ranking** | Weighted score — Price 40% · Delivery 25% · Quantity 20% · Proposal Type 15% |
| ⚠️ **Variance Analysis** | Z-score anomaly detection — flags suspiciously low bids and price outliers |
| 🔄 **Alternate Proposals** | Side-by-side original vs alternate with % cost impact and recommendation |

---

## Expected CSV Format

```csv
Bid_ID,Supplier,Category,Product,Quantity,Unit_Price,Total_Amount,Delivery_Days,Proposal_Type
BID001,Apex Materials,Packaging,Carton Box,500,4.50,2250,5,Original
BID002,SwiftPack Ltd,Packaging,Carton Box,500,4.20,2100,8,Alternate
```

> **Flexible parsing** — column names do not need to match exactly. BID-X auto-detects common aliases (`Vendor`, `Rate`, `Lead Time`, `Quote Type`, etc.). If detection fails, a column-mapping modal appears so you can assign fields manually.

---

## Tech Stack

```
Frontend   Vanilla HTML · CSS · Canvas API (animated charts)
Backend    Python 3.11 · Flask 3.1 · Pandas 2.2
Server     Gunicorn (2 workers)
Hosting    Render.com (free tier)
```

---

## Local Setup

```bash
# 1. Clone
git clone https://github.com/chaqBD/BID-X.git
cd BID-X

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
python app.py
# → http://localhost:5000
```

---

## Deploy to Render

1. Fork or push this repo to your GitHub account
2. Go to [render.com](https://render.com) → **New Web Service**
3. Connect the repo — Render reads `render.yaml` automatically
4. Click **Deploy** — live in ~2 minutes

---

## Analysis Workflow

```
Upload CSV
    │
    ▼
Data Preview  (first 10 rows shown instantly)
    │
    ▼
Choose module from sidebar  ──►  Run All Analyses
    │
    ▼
Results panel  (charts · tables · anomaly flags)
    │
    ▼
Export CSV report
```

---

## License

MIT © [X-CHAQ LTD, UK](https://github.com/chaqBD)

---

<div align="center">
  <sub>Built with care by <strong>X-CHAQ LTD, UK</strong></sub>
</div>
