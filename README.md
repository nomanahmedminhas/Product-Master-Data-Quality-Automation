# Product Master Data Quality Automation | Excel VBA

## Project Overview

An Excel-based product master data quality and governance project that automates data cleansing, duplicate detection, exception tracking and quality reporting using VBA.

The project demonstrates how raw product catalog data can be transformed into a standardized, reviewable master dataset while preserving the original source.

## Tools and Technologies

- Microsoft Excel 2010/2013
- Visual Basic for Applications (VBA)
- Excel Tables and Conditional Formatting
- COUNTIF, COUNTBLANK, VLOOKUP, EXACT and IF formulas
- Excel Charts and KPI Dashboards

## Key Features

- Automated standardization of product categories, suppliers, colors and statuses
- Duplicate SKU detection and provisional deduplication
- Missing-data validation
- Automated Exception Log
- Review Queue for unresolved issues
- Dynamic data-quality dashboard
- One-click VBA refresh
- Preservation of original source data

## Demonstration Results

The sample dataset contains 15 raw records, producing 14 unique SKU records.

The most recently reported exception summary contains:

| Metric | Result |
|---|---:|
| Raw Records | 15 |
| Unique SKUs | 14 |
| Duplicate Groups | 1 |
| Total Exceptions | 20 |
| Resolved Exceptions | 17 |
| Pending Exceptions | 3 |
| Exception Resolution Rate | 85% |

These results describe a demonstration dataset, not a production system.

## How to Use

1. Download the macro-enabled Excel workbook.
2. Open it in desktop Microsoft Excel.
3. Review the sample records in Raw Product Data.
4. Enable macros only after reviewing and trusting the VBA code.
5. Open the Dashboard and click **REFRESH DATA**.
6. Review Clean Master Data, Review Queue and Exception Log.

## Important Limitations

- Conflicting duplicate records require human verification.
- Missing values are not automatically invented.
- Clean Master Data is a provisional standardized dataset until outstanding exceptions are reviewed.
- The workbook was designed for desktop Excel with VBA support.
- This is an educational portfolio demonstration, not a production-certified MDM solution.

## Skills Demonstrated

Master Data Management (MDM), Product Information Management (PIM), Data Quality, Data Governance, Excel VBA Automation, Data Cleansing, Validation, Exception Management and Dashboard Reporting.
