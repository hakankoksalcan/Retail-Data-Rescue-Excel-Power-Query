# Validation ledger — 2026-10-21

Different evidence supports different claims. PASS applies to the named scope, not to every part of the project.

| Scope | Evidence | Status |
|---|---|---|
| Original UCI workbook integrity | Raw SHA-256 and source schema | Checked by publication validation |
| Reference classifications / row and amount conservation | `reference/Controls.csv` | Executed in Python; explicit Excel acceptance row remains pending |
| Same-source and +GBP 7 scenario | `refresh-tests.json`, `refresh-controls.csv` | Executed in Python |
| All supplied M sources, including four embedded exports | `m-syntax.json` | Microsoft parser syntax validation; no semantic/runtime claim |
| Source audit refreshed in Excel 2021 | Uploaded SourceCheck table and user screenshot/report | 541909 / 8 confirmed |
| Missing-value audit refreshed in Excel 2021 | Uploaded qMissingValues table and user report | CustomerID 135080; Description 1454; other fields 0 |
| qTypedSource | Embedded M and user preview confirmation | Present; full-table conversion-error scan not separately evidenced |
| Six filtered Excel KPI results | `user-excel-filter-confirmation.json` | All six matched, user-reported |
| Spreadsheet formula recalculation | `libreoffice-filter-check.json` | Six results checked in LibreOffice; not M execution |
| Main XLSX package / cached tables / charts | `publication-validation.json` | Fresh artifact checks; see individual results |
| Query package preservation during presentation fixes | `workbook-preservation.json` | Query, connection, table and chart parts unchanged |
| Complete reporting M model refreshing dashboard | `EXCEL_ACCEPTANCE_CHECKLIST.csv` | Pending full-model integration and desktop Excel execution |

## Workbook provenance

The user supplied an Excel-authored workbook containing four queries. The published workbook preserves its DataMashup package, connections, query tables and charts byte-for-byte. Presentation changes update obsolete status text, normalize the configurable sample path, remove local Desktop presentation metadata, and open the Dashboard by default. Values are never invented to replace failed controls.

The main Dashboard / reporting tables remain the independently generated reference snapshot; their formulas and filters work in Excel. Refresh All currently refreshes the embedded audit scope. The separate 20-query reporting model is not represented as already installed.

Microsoft Excel is unavailable in this Linux validation environment. Existing desktop Excel evidence comes from the user's uploaded workbook and explicit confirmations; it is not an automated Excel run performed here. The workbook preview is a LibreOffice render, labelled accordingly in its provenance.

## Re-run

Follow the main README. `validate_publication.py` exits nonzero for failed checks, writes individual expected/actual values, checks CSV/reference reconciliation, inspects the Excel query package, validates relative documentation links and checks the file manifest. M syntax and Python refresh checks are separate commands. The complete Excel acceptance checklist remains pending until it is actually performed.
