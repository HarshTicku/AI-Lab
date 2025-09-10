# Air Bnb Analysis using Advanced Excel

## Project Overview  
- **Purpose**: Analyze real estate listings data to identify optimal properties for investment and forecast returns.  
- **Core deliverables**: An Excel workbook with a BI dashboard and a printable PDF report.  
- **Use case**: Support property purchase decisions by estimating ROI and comparing scenarios across markets and property types.  

---

## Introduction  
This project provides a structured, repeatable workflow to ingest listing data, clean and standardize it, analyze market performance, and visualize insights in an executive-ready dashboard. It is fully customizable and free of any original author identifiers.

### Tools Used  
- Advanced formulas & functions (XLOOKUP, INDEX/MATCH, etc.)  
- Pivot Tables & Pivot Charts  
- What-If Analysis (Goal Seek)  
- Modern Charts (Sunburst, Box & Whisker, Treemap)  

### Workflow Stages  
1. Data ingestion  
2. Cleaning and processing  
3. Analysis and modeling  
4. Dashboard and reporting  

---

## Worksheet & Tab Documentation  
The workbook organizes logic by tab. Use the `Documentation` tab inside the workbook for in-sheet guidance.

- **Data_Ingestion**: Paste/import raw listings data here (e.g., CSV, exports). Required inputs include listing ID, location, bedrooms, price, occupancy, fees.  
- **Data_Cleaning**: Click the `Run Data_Cleaning` macro to standardize headers, remove duplicates, normalize text/date formats, handle missing values, and apply lookups.  
- **Analysis**: Computed fields for ADR, RevPAR, net revenue, cost assumptions, and ROI/Payback. Contains staging tables for pivots.  
- **Dashboard**: Interactive visuals connected to analysis pivots; includes slicers for City, Property Type, Bedrooms, Time.  
- **Documentation**: Tab-level instructions, data map, formula glossary, refresh checklist.  

Data flows from `Data_Ingestion` → `Data_Cleaning` → `Analysis` → `Dashboard`.

---

## Quickstart  
1. Open `Real Estate Investment Dashboard.xlsx` and enable macros when prompted.  
2. Load sample data: use `Data_Ingestion` → Paste from `Sample Listings Dataset.xlsx` (or import CSV).  
3. Run cleaning: Developer → Macros → `Run_Data_Cleaning` (or click the in-sheet button).  
4. Refresh pivots and charts: Data → Refresh All.  
5. Explore the dashboard using slicers (City, Property Type, Bedrooms, Month).  

> Tip: Right-click any pivot → Refresh if individual visuals are out-of-date.

---

## Dashboard Guide  
- **Sunburst**: Portfolio composition by City → Property Type → Bedroom count. Insight: exposure and diversification.  
- **Box & Whisker**: Distribution of ADR/revenue by segment. Insight: variability and outliers.  
- **Treemap**: Contribution to total revenue by segment. Insight: key drivers.  
- **Line/Bar**: Trend of occupancy, ADR, and RevPAR over time. Insight: seasonality and performance.  

---

## Extending the Project  
- Add new fields: Append columns in `Data_Ingestion`, map them in `Data_Cleaning`, and add to `Analysis` staging tables; Refresh All.  
- Integrate new sources: Import CSV/DB extracts into `Data_Ingestion` and ensure headers match; use Power Query or manual import.  
- Scenario changes: Modify `Analysis` assumptions (costs, occupancy, price) to test different outcomes.

---

## License & Attribution  
This project is provided as a neutral, customizable template for real estate analytics. All prior author identifiers have been removed.



