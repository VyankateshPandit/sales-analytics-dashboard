# Sales Analytics Dashboard with Data Quality Layer

A Python-based sales analytics pipeline that generates a fully structured Excel workbook with a built-in data quality audit layer — ready for Power BI connection.

---

## Tech Stack

`Python` `Pandas` `openpyxl` `Excel` `Power BI`

---

## Project Structure

```
├── sales_dashboard.py              # Main script — generates the full workbook
├── Sales_Analytics_Dashboard.xlsx  # Output workbook (4 sheets)
└── README.md
```

---

## What It Does

### 1. Synthetic Dataset Generation
- 500+ sales orders across 50 customers, 8 products, 5 regions
- Realistic pricing, discounts, order statuses, and date ranges
- Intentionally injected data issues: missing values, invalid orders, duplicates

### 2. Data Quality Layer (Sheet: `Data_Quality`)
Automatically detects and reports:

| Issue Type | Detection Method | Severity |
|---|---|---|
| Missing values | Null check on Customer ID, Price, Region | HIGH |
| Invalid orders | Negative / zero quantity or price | HIGH / MEDIUM |
| Duplicate transactions | Exact row + Order ID deduplication | MEDIUM |

- Flags ~15% of raw records as dirty
- Provides business impact and remediation action per issue
- Generates a quality score (clean records / total records)

### 3. Sales Dashboard (Sheet: `Sales_Dashboard`)
Built exclusively on validated clean data:
- Monthly revenue trend (line chart)
- Top 10 customers by revenue
- Regional performance with revenue share
- Product breakdown by revenue and units sold

### 4. Clean Data Export (Sheet: `Clean_Data_PowerBI`)
- Post-validation dataset ready for direct Power BI import
- Connect via: Power BI Desktop → Get Data → Excel → `Clean_Data_PowerBI`

---

## How to Run

```bash
pip install pandas openpyxl
python sales_dashboard.py
```

Output: `Sales_Analytics_Dashboard.xlsx`

---

## Key Results

- **515** raw records generated
- **438** clean records after quality filtering
- **84.5%** data quality score
- **4-sheet** workbook: Raw Data → Quality Report → Dashboard → Clean Export
