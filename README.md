# Sales Performance MIS Dashboard (Excel + VBA)

An interactive Excel dashboard that tracks the sales performance of **142 sales executives across 8 regions**. It shows Top 5 / Bottom 5 performers, target achievement and gap-to-target, and lets the user choose which dashboard section updates when a region is selected.

<img width="1250" height="870" alt="dashboard" src="https://github.com/user-attachments/assets/fb67ddfe-a7f0-495f-9590-0e1ab291fc02" />

---

## Problem Statement

Management needs to quickly see:
- Which sales executives are performing best and worst
- How close each executive is to the monthly target
- How each region is performing

Doing this manually from raw data is slow and error-prone. This dashboard gives the answers in one click.

---

## Dataset

| Item | Details |
|---|---|
| Records | 142 sales executives |
| Regions | Chennai, Delhi, Mumbai, Nagpur, Patna, Pune, Ranchi, Surat |
| Period | 5 days of sales (Day1 to Day5) |
| Target | 500 units per executive |

**Columns:** Emp Code, Sales Executive, Region, Day1, Day2, Day3, Day4, Day5, Total Sales, Target, Target Hit %, Away From Target %

---

## Calculated Columns

| Column | Formula |
|---|---|
| Total Sales | `=SUM(D2:H2)` |
| Target Hit % | `=I2/J2` |
| Away From Target % | `=1-K2` |

---

## Dashboard Features

- **Region Slicer:** 8 region buttons to filter the data
- **4 Pivot Table sections** (Top 5 and Bottom 5 by Total Sales, Target Hit %, Away From Target %)
- **3 Pivot Charts:** bar chart, pie chart, line chart
- **Selective filtering:** 4 checkboxes (Dashboard 1 to 4). Tick a checkbox, then click a region, and only that dashboard section updates. The others stay unchanged.

---

## How to Use

1. Tick the checkbox of the dashboard section you want to update (Dashboard 1 / 2 / 3 / 4).
2. Click a region button (for example, Mumbai).
3. The tables and charts of the selected section update for that region.
4. Untick the checkbox to stop that section from responding to the slicer.

---

## How the Selective Filter Works (VBA)

A normal slicer filters all connected pivot tables at once. To control this, each checkbox is linked to a cell (Cell Link), which becomes `TRUE` when ticked and `FALSE` when unticked. A macro reads that cell and connects or disconnects the Region slicer from the matching pivot table.

```vba
Sub SlicerConnection()

    ' Checkbox linked cell A1: TRUE = connect, FALSE = disconnect
    If Sheet1.Range("A1").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable Sheet1.PivotTables("PivotTable1")
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable Sheet1.PivotTables("PivotTable1")
    End If

End Sub
```

The same logic is repeated for the other dashboard sections with their own linked cell and pivot table.

---

## Tools and Skills Used

- Microsoft Excel (Formulas, Pivot Tables, Pivot Charts)
- Slicers and Slicer Caches
- Form Controls (Checkboxes, Cell Link)
- VBA (Macros, If...Else, SlicerCaches, PivotTables objects)
- Conditional logic and dashboard design

---

## Project Structure

```
Sales-MIS-Dashboard/
|-- Sales_MIS_Dashboard.xlsm      (main file: Sheet1 = DASHBOARD)
|-- screenshots/
|   |-- dashboard.png             (full dashboard view)
|   |-- dashboard_region.png      (dashboard after selecting a region)
|-- vba/
|   |-- SlicerConnection.bas      (exported VBA macro code)
|-- data/
|   |-- raw_data.csv              (RAW DATA sheet)
|-- README.md
```

The dashboard and the raw data live in **one Excel workbook** (two sheets). The `vba`, `screenshots` and `data` folders are only for easy viewing on GitHub.

> Note: Download the **.xlsm** file and click **Enable Content** when Excel asks, otherwise the macros and checkboxes will not work.

---

## Key Learnings

- Building interactive dashboards with Pivot Tables, Charts and Slicers
- Controlling slicer connections with VBA
- Calculating KPIs such as Target Hit % and Gap-to-Target
- Presenting data so that management can act on it quickly

---

## Future Improvements

- Add KPI cards (Total Sales, Average Target Hit %, Total Executives)
- Add titles to every chart and table
- Add a Month dimension for trend analysis
- Rebuild the same dashboard in Power BI

---

## Author

**[Harsh Kushwah]**
Email: [harshuu9752@gmail.com]
