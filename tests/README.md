# Reference repeat / appended-source test

Run `python scripts/test_refresh.py` after the reference generation steps in the main README.

The script reads the original source again and compares totals, class partitions and nine output tables. It then writes a separate full-size `TEST_ONLY_Online_Retail_plus_one.xlsx` and appends this labelled fixture:

| Field | Value |
|---|---|
| InvoiceNo | 990001 |
| StockCode | 85123A |
| Description | WHITE HANGING HEART T-LIGHT HOLDER |
| Quantity | 2 |
| InvoiceDate | 2011-12-09 13:00 |
| UnitPrice | 3.5 |
| CustomerID | 17850 |
| Country | United Kingdom |

Expected differences: one source row, one eligible sale invoice, GBP 7 additional raw/sales/balance amount; cancellations and duplicate-extra count unchanged. New SourceRow: 541911. The original workbook's checksum must remain unchanged.

The synthetic workbook is generated locally, excluded from Git, and never substituted for the original source. `repeat_reference/` and `appended_reference/` contain comparison tables. Results are in `validation/refresh-tests.json` and `refresh-controls.csv`.

These are Python reference tests. Repeating them through the full Power Query model in desktop Excel is a separate, pending acceptance step.
