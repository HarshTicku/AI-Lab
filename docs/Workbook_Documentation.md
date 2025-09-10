# Workbook Documentation

This mirrors the in-workbook `Documentation` tab.

## Tabs & Purpose
- Data_Ingestion: Paste/import raw listings data. Required columns: ListingID, City, Zip, Bedrooms, PropertyType, NightlyPrice, CleaningFee, OccupancyRate, Date.
- Data_Cleaning: Output of the `Run_Data_Cleaning` macro. Standardized headers, types, and enriched fields.
- Analysis: Calculations for ADR, RevPAR, Net Revenue, Costs, Profit, ROI, Payback; staging tables for pivots.
- Dashboard: Pivot charts, KPI cards, slicers.
- Documentation: Instructions, data dictionary, refresh checklist.

## Inputs
- Paste CSV/Excel extracts into `Data_Ingestion` starting at A1, including headers.
- Keep consistent header names to maintain formula references.

## Key Formulas
- ADR: =IFERROR(NetRevenue/NightsOccupied,0)
- RevPAR: =ADR*OccupancyRate
- NetRevenue: =NightlyPrice*NightsOccupied - CleaningFee - OtherCosts
- ROI: =IFERROR(AnnualProfit/Investment,0)

## Data Flow
Data_Ingestion → Data_Cleaning → Analysis → Dashboard

## Refresh Checklist
1) Run `Run_Data_Cleaning` macro  
2) Data → Refresh All  
3) Validate slicers and timelines  
