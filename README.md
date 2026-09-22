# CFM Quote Data Tool

An offline-capable, standalone quotation workbench and pricing matrix application for CFM Distributors HVACR commercial equipment, residential split systems, packaged rooftop units (RTUs), and paired accessories.

![CFM Quote Data Tool](cfm_logo.png)

---

## Live Application

The interactive web application is deployed and hosted on GitHub Pages:
👉 **[https://jraltd.github.io/cfm-Quote-Data-Tool/](https://jraltd.github.io/cfm-Quote-Data-Tool/)**

---

## Overview

The **CFM Quote Data Tool** (`cfm_quote_tool.html`) compiles equipment catalogs, pricing schedules, and accessory compatibility tables into a single portable application that runs locally in any web browser with zero external dependencies.

### Core Capabilities

- **Catalog Search & Filtering**:
  - Sub-millisecond instant search across model numbers, descriptions, and equipment families.
  - Dedicated tabs for Commercial RTUs, Residential Split Systems, Commercial Splits, Ductless & VRF, and Applied Equipment.
  - Quick filters for new low-GWP **R-454B (A2L)** systems vs. legacy **R-410A**, plus stock status indicators.

- **Paired Accessory Engine**:
  - Comprehensive accessory schedules mapped across 14 equipment series (York Sunline, Sun Pro, SunCore, SunChoice, Sun Select, Sun Premier, LX Series, Commercial Splits, Furnaces, Reznor, PTACs, Boilers, and LG Ductless).
  - One-click `⚙️ Accessories` modal popover with category filtering (Curbs, Economizers, Hail Guards, Electrical, Sensors).
  - Smart recommendations on the quote proposal whenever equipment units are selected.

- **IPA Pricing Matrices & Interactive Calculator**:
  - Multiplier tables for In-Stock (Schedule IPA 753978) and Factory Order equipment across contractor discount tiers (8% to 21%).
  - Real-time margin calculator computing Buy Cost, Target Sell Price, Dollar Margin, and Gross Margin %.

- **Contractor Proposal Builder**:
  - Customer-ready quotation builder with contractor details, project info, and automated freight calculation based on CFM policy tiers.
  - One-click Print/PDF generation with print stylesheets (`@media print`) that format clean black-on-white documents.
  - CSV export for spreadsheet records.

- **Theme Engine**:
  - Dual Dark/Light themes with instant header toggle button (`☀️ Light Mode` / `🌙 Dark Mode`).
  - Automatic OS preference detection (`prefers-color-scheme: dark`) and `localStorage` persistence.

---

## File Structure

- **`cfm_quote_tool.html`**: The compiled standalone web application (~2.5 MB). No web server, node, or python runtime required for end users.
- **`generate_app.py`**: Python compiler script that extracts equipment data, matrices, and accessories from the source workbook, structures the JSON payload, and builds the HTML distribution.
- **`cfm_logo.png`**: Official CFM Distributors brand asset.
- **`Quote Data For Release 092126.xlsx`**: Source workbook containing equipment catalogs, pricing schedules, and accessory tables.

---

## Rebuilding the Application

To rebuild `cfm_quote_tool.html` from a refreshed source workbook:

```bash
# Requires Python 3.8+ and openpyxl
python generate_app.py
```

The script will process catalog worksheets, clean and structure accessory pairings, embed the brand logo as Base64, and output the standalone `cfm_quote_tool.html` file.
