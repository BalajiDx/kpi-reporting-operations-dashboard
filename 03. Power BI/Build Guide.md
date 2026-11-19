# Power BI Build Guide

Steps to rebuild the report in Power BI Desktop from this repo. The `.pbix` itself is not included; save yours as
`03. Power BI/File/Operations_KPI_Dashboard.pbix` once built. The images in `Dashboard Images/` show the target layout.

## 1. Load the data
1. **Home > Get data > Text/CSV**, and load these files from `01. Dataset/Cleaned Dataset/`:
   - `Clean_Jobs.csv` -> rename the query to `Jobs`
   - `Clean_Staffing.csv` -> rename the query to `Staffing`
   - `Clean_Escalations.csv` -> rename the query to `Escalations`
2. In Power Query set types: dates to **Date**, money and rates to **Decimal number**, counts to **Whole number**, IDs to **Text**.
3. Create `Teams` and `Calendar` from `DAX/Measures.dax`.

## 2. Model
Create these relationships (single direction, filter flows from the "1" side):

| From (1) | To (*) |
|---|---|
| Calendar[Date] | Jobs[DeliveredDate] |
| Calendar[Date] | Jobs[ReceivedDate] (inactive) |
| Calendar[Date] | Escalations[Date] |
| Teams[Team] | Jobs[Team] |
| Teams[Team] | Staffing[Team] |
| Teams[Team] | Escalations[Team] |

Mark `Calendar` as the date table.

## 3. Add DAX
From `DAX/Measures.dax`, in order: the `Calendar` table, calculated columns (set **Sort by column** on band columns),
then measures in a dedicated `_Measures` table with display folders (Delivery, Quality, Staffing).

## 4. Pages
| Page | Image | Visuals |
|---|---|---|
| 1. KPI Overview | `1.KPI_Overview.png` | KPI cards, on-time line with 90% target line, delivered columns, team lines, last-six-months table |
| 2. Team Scorecard | `2.Team_Scorecard.png` | Month slicer, team table with conditional formatting on status, staffing line, escalations columns |
| Model | `3.Data_Model.png` | Relationships to build in Model view |

## 5. Check your numbers
With no slicers applied the report should show:
- On-time % ≈ 83.0%; First-pass Quality % ≈ 97.1%
- Oct-24: On-time % ≈ 66.2%, Typesetting ≈ 53.3%
- Every monthly figure should match the Excel 'Monthly KPIs' sheet

The full set of reference results is in `02. SQL Scripts/Query Results/`. Anything different means a relationship or a data type is wrong.

## 6. Optional: Power BI Service
Publish to a workspace, set **scheduled refresh** (a gateway is needed for on-premises sources), add **row-level security**
under Modeling > Manage roles, and use a **deployment pipeline** to move the report from Development to Test to Production.
