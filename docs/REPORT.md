# Real Estate Investment Dashboard — Report

## Cover
- Title: Real Estate Investment Dashboard
- Author: Data Analyst
- Organization: Acme Analytics
- Version: 1.0.0
- Date: YYYY-MM-DD

## Executive Summary
This report summarizes key insights from the Real Estate Investment Dashboard built in Excel. It identifies attractive properties and markets, estimates potential returns, and highlights the drivers of portfolio performance.

## Methodology
- Data sources: Listings exports (CSV/Excel) with attributes such as location, bedrooms, nightly price, occupancy, and fees.
- Processing: Standardized headers, removed duplicates, normalized data types, handled missing values, and mapped categories.
- Analysis: Calculated ADR, RevPAR, net revenue, operating costs, profit, and ROI. Conducted scenario analysis and what-if testing.
- Visualization: Excel Pivot Tables/Charts with slicers; Sunburst, Box & Whisker, Treemap, and trend visuals.

## Key Insights
- Top markets by net revenue contribution and risk (variability).
- Property types and bedroom counts with favorable ROI and shorter payback.
- Seasonality patterns impacting occupancy and price.

## Dashboard Guide
Refer to `docs/Dashboard_Guide.md` for a visual guide and business interpretation of each chart.

## Version History
- 1.0.0 — Initial neutralized release, metadata sanitized, README updated, and macros consolidated.

## Appendix A — Next Steps
- Add new data fields: Extend `Data_Ingestion` with new columns and map them in `Data_Cleaning` and `Analysis` staging tables. Refresh All.
- Integrate additional data sources: Import CSVs or database extracts (e.g., ODBC/Power Query) into `Data_Ingestion` with consistent headers.
- Modify forecasting: Adjust analysis assumptions (occupancy, pricing, costs) to evaluate alternate scenarios.

## Appendix B — Metadata & Branding Sanitation
Ensure the workbook and exported PDFs are free of personal identifiers. See `docs/metadata_sanitization.md` for step-by-step instructions.


