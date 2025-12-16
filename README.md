# Excel Monthly Financial Combiner

A Python utility that scans a `YYYY/MM.YYYY` folder structure, reads a monthly Excel workbook in each month folder, and generates a single combined Excel file with three sheets:

- `P&L Combined`
- `BS Condensed Combined`
- `DataBase Combined`

It is designed for pipelines where a dashboard or BI tool consumes a consolidated workbook, while source files are stored per month.

## Folder layout expected

