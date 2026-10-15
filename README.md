# 📊 Sales & Inventory Performance Dashboard
An interactive Excel dashboard that tracks **revenue, profit, sales trends, and stock health** for a 30-product catalog across four regions and three sales channels. Raw transaction and product data was cleaned with **Power Query**, enriched with formulas (`XLOOKUP`, `SUMIFS`, `IFS`), and analyzed with **PivotTables, PivotCharts, a Slicer, and a Timeline**.


## 🎯 Business Problem
A multi-channel retailer needs to answer four questions from its sales and inventory records:

1. How much revenue and profit are we generating, and where does it come from (channel, region, category)?
2. How is revenue trending month over month?
3. Which products are healthy, which need reordering, and which are already in a stock-out position?
4. Can stakeholders filter all of this themselves without touching the data?


## 🔑 Key Findings
| Metric | Value |
|---|---|
| Total revenue | **572.6K** (850 transactions, Jan 1 – Oct 7, 2026) |
| Total units sold | **6,370** |
| Gross profit | **354.8K** (≈ **62%** margin) |
| Top sales channel | **Online** – 200.2K (35% of revenue), ahead of Wholesale (192.5K) and In-Store (179.8K) |
| Top region | **North America** – 153.7K, followed by Asia-Pacific (148.8K) |
| Top category by profit | **Electronics** – 178.5K (≈ 50% of total profit) |
| Best-selling product by revenue | **Ultrawide 34in Monitor** – 94.3K (16% of revenue) |
| Peak month | **April 2026** – 83.6K revenue |

### ⚠️ Inventory Risk
- **19 of 30 products (63%) are in STOCK CRISIS**, 4 need REORDER, and only 7 are STABLE.
- Total units sold (6,370) exceed total starting stock (5,680), leaving a **cumulative shortfall of 2,120 units** across the 19 crisis products.
- Those 19 products generate roughly **88% of revenue and profit**, so the stock-out risk sits on the revenue-critical items.

---

## 🛠️ Tools & Skills Used
- **Microsoft Excel**: PivotTables, PivotCharts, Slicer, Timeline, conditional formatting
- **Power Query**: data cleaning and transformation
- **Excel formulas**: `XLOOKUP`, `SUMIFS`, `IFS`, `IF`, `COUNTIF`, `DATE`, `TEXT`, `DATEVALUE`
- **Dashboard design**: KPI cards, interactive filtering, chart-based storytelling

---

## 🧹 Data Cleaning (Power Query)
| Source sheet | Rows | Issues found | Fix |
|---|---|---|---|
| `Raw_Sales_Data` | 850 | 91 Product IDs in inconsistent case (e.g. `prod-002`) | Standardized to upper case (`PROD-002`) |
| | | Region values with trailing spaces (e.g. `"Europe "`) | Trimmed whitespace so each region forms one group |
| | | Two date formats mixed in one column (`2026-04-13` and `07/21/2026`) | Converted to a single true date (`True_Date`) |
| `Raw_Inventory_Master` | 30 | Inconsistent product-name casing | Standardized to proper case |

After cleaning: **no duplicate Transaction IDs, no missing values, and every sales row maps to a valid product.**

---

## 🧮 Data Model & Calculations

### Sales sheet: `Clean_Sales_Data` (850 rows)

| Column | Logic |
|---|---|
| `True_Date` | Parses both ISO (`YYYY-MM-DD`) and `MM/DD/YYYY` text dates into a real Excel date |
| `REVENUE` | `Quantity Sold × XLOOKUP(Product ID → Selling Price)` |
| `Month` | `DATE(YEAR, MONTH, 1)`: a real date used to sort months chronologically |
| `Month field` | `TEXT(Month, "mmm yyyy")`: readable label for charts |

### Inventory sheet: `Clean_Inventry_Master` (30 rows)

| Column | Logic |
|---|---|
| `Total Unit Sold` | `SUMIFS(Quantity Sold, Product ID = this product)` |
| `Ending Inventory` | `Starting Stock − Total Unit Sold` |
| `Gross Revenue` | `Total Unit Sold × Selling Price` |
| `Profit` | `Gross Revenue − (Total Unit Sold × Unit Cost)` |
| `Inventory Status` | `IFS`: **STOCK CRISIS** if Ending Inventory < 0; **REORDER** if ≤ Safety Stock Level; otherwise **STABLE** |

---

## 📈 Dashboard Components

| Visual | Purpose |
|---|---|
| **KPI cards** | Total Revenue, Total Units Sold, Gross Profit at a glance |
| **Revenue by Sales Channel** (bar) | Compares Online, Wholesale, and In-Store |
| **Profit by Category** | Shows which category drives margin |
| **Monthly Revenue Trend** (line) | Tracks seasonality and month-over-month movement |
| **Total Products by Inventory Status** (clustered column) | Counts products in REORDER / STABLE / STOCK CRISIS by category |
| **Region Slicer** | Filters every chart by Asia-Pacific, Europe, Latin America, or North America |
| **Date Timeline** | Filters every chart by month or period |

---

## 📁 Repository Structure

```
├── Enterprise_Inventory_and_Sales_Data.xlsx   # Full workbook (raw + clean data, pivots, dashboard)
├── images/
│   └── dashboard.png                          # Dashboard screenshot
└── README.md
```

**Workbook sheets**

| Sheet | Contents |
|---|---|
| `Dashboard` | Final interactive dashboard |
| `Clean_Sales_Data` | Cleaned transactions + calculated columns |
| `Clean_Inventry_Master` | Cleaned product master + inventory calculations |
| `Raw_Sales_Data` / `Raw_Inventory_Master` | Original untouched data |
| `chart`, `monthly revenue`, `sales channel revenue` | Pivot tables feeding the dashboard charts |

---

## ▶️ How to Use

1. Download `Enterprise_Inventory_and_Sales_Data.xlsx`.
2. Open it in **Microsoft Excel (Microsoft 365 or Excel 2021+)**. `XLOOKUP` and `IFS` require a recent version.
3. Go to the **Dashboard** sheet and use the **Region slicer** and **Date timeline** to explore the data.
4. To refresh after editing the data: **Data → Refresh All**.

---

## 💡 Recommendations

1. **Prioritize replenishment of high-revenue crisis items** such as the Ultrawide 34in Monitor, Electric Standing Desk, and Noise Cancelling Headphones, which together drive a large share of profit.
2. **Re-set safety stock levels**: most products sold more than their opening stock, so current safety levels are too low for actual demand.
3. **Invest in the Online channel and Electronics category**, the largest revenue and profit contributors.
4. **Investigate the Jan, Mar, and May revenue dips** (≈ 48–54K vs. 66K+ in Feb, Jul, Aug, and Sep) to see whether they reflect seasonality or supply constraints.

---

## ⚠️ Assumptions & Limitations

- The dataset has **no restocking or purchase-order data**, so Ending Inventory = Starting Stock − Units Sold. Negative values represent **unmet demand / backorders**, not physical stock.
- **October 2026 is a partial month** (data ends Oct 7), so its revenue should not be compared directly with full months.
- Profit is calculated as **gross profit** (selling price − unit cost); operating expenses, discounts, and returns are not included.

---

## 🚀 Future Improvements

- Add restock / purchase-order data to model true stock levels and lead times.
- Add reorder-point and days-of-stock-remaining calculations.
- Rebuild the dashboard in Power BI for sharing and scheduled refresh.

---

## 👤 Author

**yogita**
