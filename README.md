# Student Performance Analysis & Interactive Dashboard

**Student:** Priyanshu Buhecha  
**Course:** AI / ML — Excel Data Analysis Practical  
**Workbook:** `Priyanshu_Buhecha_Student_Data_Dashboard.xlsx`  
**Dataset:** Student performance with demographic, study and assessment attributes.

## Project objective
Import student performance data, clean and prepare it, summarize it with pivot tables and formulas, visualize patterns, build a dashboard, and document the process.

## Dataset summary
- **Valid score records analyzed:** 30,641
- **Average mathematics score:** 66.56
- **Average reading score:** 69.38
- **Average writing score:** 68.42
- **Key dimensions:** Gender, EthnicGroup, ParentEduc, LunchType, TestPrep, ParentMaritalStatus, PracticeSport, IsFirstChild, NrSiblings, TransportMeans and WklyStudyHours.
- These summary values are calculated from the original **StudentData** worksheet; missing numeric scores are excluded from the corresponding averages.

## Workbook structure
| Worksheet | Purpose |
| --- | --- |
| StudentData | Original student records and input fields |
| Task 1 | Dataset selection, import and first checks |
| Task 2 | Data cleaning, duplicate detection and missing-value review |
| Task 3 | Summarization and source for pivot-style analysis |
| Pivot Table (Task 3) | Pivot-table workspace |
| Task 4 | Advanced lookup and conditional formula practice |
| Task 5 | Charts and conditional-formatting exercise |
| Task 6 | Interactive dashboard and dynamic range exercise |
| README, Sheet Index, Data Dictionary, Task Log, Instructions, Naming Rules, Quality Checks | Workbook documentation and usage guidance |

## Task-by-task explanation

### Task 1 — Dataset selection and import
Load the student performance dataset into Excel, verify field headers, inspect score columns, and retain an untouched raw-data sheet. **Output:** imported student records in `StudentData` and `Task 1`.

### Task 2 — Data cleaning and preparation
Review duplicates, blanks and numeric data types. Excel functions can include `COUNTIF`, `COUNTBLANK`, `IF`, `TRIM`, and `VALUE`. **Output:** cleaning exercise in `Task 2`; check whether detected issues were actually resolved before presenting a clean dataset.

### Task 3 — Pivot tables and summaries
Summarize counts by gender or test preparation and compare subject averages. Insert PivotTable using the data range, put a category in Rows and a score in Values, then change Values to Average as needed. **Output:** `Task 3` and `Pivot Table (Task 3)`.

### Task 4 — Advanced formulas
Practice `VLOOKUP`, `INDEX`/`MATCH`, `IF`, nested `IF`, and `IFERROR`. Check Excel version and locale before using dynamic-array functions. **Output:** `Task 4`.

### Task 5 — Data visualization
Create column charts for average subject scores, category counts and test-preparation comparisons. Apply conditional formatting to flag score bands. **Output:** `Task 5` and the analytical PNG charts in this package.

### Task 6 — Interactive dashboard
Use PivotCharts, slicers, and (where appropriate) a timeline for actual date fields. **This dataset has no obvious date field, so a timeline requires adding a legitimate date dimension.** Dynamic-array functions such as `FILTER` may show `#NAME?` in older Excel versions; use PivotTable filters or compatible helper columns instead. **Output:** `Task 6`, plus a separately reconstructed dashboard overview image in this package.

### Task 7 — Documentation
Maintain a sheet directory, data dictionary, task log, quality checks and user instructions. **Output:** workbook documentation sheets and this README.

## Output pictures

### Dashboard overview
![Dashboard overview](screenshots/01_dashboard_overview.png)

### Average score chart
![Average scores](screenshots/02_subject_score_chart.png)

### Gender distribution
![Gender distribution](screenshots/03_gender_distribution.png)

### Test-preparation comparison
![Test preparation](screenshots/04_test_prep_comparison.png)

### Worksheet previews
![Task 3 cell preview](screenshots/05_task3_worksheet_preview.png)

![Task 6 cell preview](screenshots/06_task6_worksheet_preview.png)

> **Screenshot accuracy note:** PNGs 01–04 are analytical visuals recreated from workbook data, not screenshots of the native Excel dashboard. PNGs 05–06 are text previews of actual worksheet cell contents, not screenshots of Excel's interface. The original workbook is included unchanged; verify any native charts, slicers, and formula results by opening it in Microsoft Excel.

## How to use
1. Open the `.xlsx` workbook in Microsoft Excel (desktop recommended).
2. Enable editing if prompted; inspect `README` and `Sheet Index`.
3. Visit Tasks 1–6 in sequence and verify pivot tables and chart connections.
4. Refresh pivot tables via **Data → Refresh All** after changing data.
5. Open the `screenshots/` folder to review the accompanying output illustrations.
6. When submitting, include the original workbook, README and screenshot folder.

## Limitations
- The supplied workbook was preserved rather than modified.
- A recreated dashboard PNG is not proof that slicers or charts work interactively inside Excel.
- Some formulas and Excel-only visual objects may not render in third-party spreadsheet readers.
- Data-quality and formula outputs should be verified in Microsoft Excel before grading.
