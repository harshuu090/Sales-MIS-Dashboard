<div align="center">

# 📊 Sales Performance MIS Dashboard

<img src="https://readme-typing-svg.demolab.com?font=Fira+Code&weight=600&size=24&pause=1000&color=36BCF7&center=true&vCenter=true&width=700&lines=Interactive+Sales+MIS+Dashboard;Built+with+Excel+%26+VBA;Track+Performance+%7C+Analyze+Targets;Turning+Sales+Data+into+Insights" alt="Animated Typing SVG" />

**An interactive Excel dashboard to analyze sales performance across 8 regions and 142 sales executives.**

<br/>

![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/Automation-VBA-blue?style=for-the-badge)
![Dashboard](https://img.shields.io/badge/Project-MIS%20Dashboard-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)

</div>

---

## 📸 Dashboard Preview

<div align="center">

<img src="https://github.com/user-attachments/assets/fb67ddfe-a7f0-495f-9590-0e1ab291fc02" alt="Sales Performance MIS Dashboard" width="90%"/>

</div>

---

## 🎯 Project Overview

This project focuses on analyzing sales performance using **Microsoft Excel, Pivot Tables, Pivot Charts, Slicers, Form Controls, and VBA macros**.

The dashboard helps management quickly identify top and bottom performers, monitor target achievement, and compare regional sales performance. Its selective filtering functionality allows users to update a chosen dashboard section without changing the others.

## ❗ Problem Statement

Analyzing sales performance manually can be time-consuming and error-prone. Management needs a simple way to:

- Identify the top and bottom sales executives.
- Track sales against a target of 500 units per executive.
- Calculate target achievement and the remaining target gap.
- Compare performance across different regions.
- Filter individual dashboard sections independently.

**Solution:** An interactive Excel MIS dashboard with region-based filtering and VBA-powered control over Pivot Table connections.

---

## 📁 Dataset Overview

| Metric | Details |
|---|---|
| Sales Executives | 142 |
| Regions | 8 |
| Sales Period | 5 days |
| Individual Target | 500 units |
| Data Source | Excel workbook |
| Main Worksheet | `RAW DATA` |
| Dashboard Worksheet | `DASHBOARD` |

### 🌍 Regions Covered

`Chennai` · `Delhi` · `Mumbai` · `Nagpur` · `Patna` · `Pune` · `Ranchi` · `Surat`

### 📋 Dataset Columns

`Emp Code` · `Sales Executive` · `Region` · `Day1` · `Day2` · `Day3` · `Day4` · `Day5` · `Total Sales` · `Target` · `Target Hit %` · `Away From Target %`

---

## 🧮 Calculated Columns

| Metric | Excel Formula | Purpose |
|---|---|---|
| Total Sales | `=SUM(D2:H2)` | Calculates total sales over five days |
| Target Hit % | `=I2/J2` | Measures target achievement |
| Away From Target % | `=1-K2` | Measures the percentage remaining to target |

*Format the percentage columns as Percentage in Excel.*

---

## ✨ Key Features

### 🏆 Top & Bottom Performers
Identify the top 5 and bottom 5 sales executives using sales and target-related performance metrics.

### 🎯 Target Achievement Analysis
Compare actual sales with the target of 500 units and evaluate target achievement and the remaining gap.

### 🌎 Region-Based Filtering
Use region buttons to analyze performance for Chennai, Delhi, Mumbai, Nagpur, Patna, Pune, Ranchi, and Surat.

### 📊 Interactive Visualizations
Includes Pivot Charts such as bar, pie, and line charts to present sales information visually.

### ⚙️ Selective Dashboard Filtering
Four checkboxes control which dashboard sections respond to region selection. The selected section updates while other sections remain unchanged.

### 🤖 VBA Automation
VBA macros manage connections between the Region slicer and the relevant Pivot Tables based on checkbox-linked cells.

---

## 🔄 How to Use the Dashboard

1. Open `Sales_MIS_Dashboard.xlsm` in Microsoft Excel.
2. Enable macros/content if prompted and if you trust the workbook.
3. Tick the checkbox for the dashboard section you want to update.
4. Select a region using the Region slicer.
5. Review the updated tables and charts for the selected region.
6. Select a different checkbox or region to explore other results.

> **Important:** Download the `.xlsm` workbook to use the interactive Excel dashboard. The VBA-powered functionality requires a desktop version of Excel with macros enabled.

---

## 💻 How Selective Filtering Works

A standard slicer can filter all connected Pivot Tables simultaneously. This project uses checkbox-linked cells and VBA to control which Pivot Table is connected to the Region slicer.

```vba
Sub SlicerConnection()

    ' Checkbox linked to cell A1
    ' TRUE = connect the Pivot Table
    ' FALSE = disconnect the Pivot Table

    If Sheet1.Range("A1").Value = True Then

        ActiveWorkbook.SlicerCaches("Slicer_Region") _
            .PivotTables.AddPivotTable Sheet1.PivotTables("PivotTable1")

    Else

        ActiveWorkbook.SlicerCaches("Slicer_Region") _
            .PivotTables.RemovePivotTable Sheet1.PivotTables("PivotTable1")

    End If

End Sub
```

The same approach is applied to the other dashboard sections using their respective checkbox-linked cells and Pivot Tables.

*Note: The example illustrates the connection logic. The slicer cache name, worksheet reference, and Pivot Table names must match the actual workbook.*

---

## 🛠️ Tools & Technologies

<div align="center">

![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=flat-square&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Automation-blue?style=flat-square)
![Pivot Tables](https://img.shields.io/badge/Pivot-Tables-orange?style=flat-square)
![Pivot Charts](https://img.shields.io/badge/Pivot-Charts-purple?style=flat-square)
![Data Analysis](https://img.shields.io/badge/Data-Analysis-informational?style=flat-square)

</div>

- **Microsoft Excel:** Data analysis and dashboard development
- **Excel Formulas:** KPI and target calculations
- **Pivot Tables:** Summarization and performance ranking
- **Pivot Charts:** Visual representation of sales data
- **Slicers:** Interactive region selection
- **Form Controls:** Checkbox-based dashboard selection
- **VBA:** Automation of slicer connections

---

## 📂 Project Structure

```text
Sales-MIS-Dashboard/
│
├── Sales_MIS_Dashboard.xlsm
├── README.md
│
├── screenshots/
│   ├── dashboard.png
│   └── dashboard_region.png
│
├── vba/
│   └── SlicerConnection.bas
│
└── data/
    └── raw_data.csv
```

The Excel workbook contains the dashboard and raw data worksheets. The additional folders organize the screenshots, exported VBA code, and raw dataset for easier review on GitHub.

---

## 📚 Key Learnings

- Building interactive MIS dashboards in Excel.
- Using Pivot Tables and Pivot Charts for performance analysis.
- Calculating target achievement and target gaps.
- Implementing region-based filtering.
- Automating Excel dashboard behavior with VBA.
- Designing dashboards that support management decision-making.

---

## 🚀 Future Improvements

- Add KPI cards for total sales, average target achievement, and total executives.
- Improve chart and table titles for easier interpretation.
- Add monthly sales data for trend analysis.
- Recreate the dashboard in Power BI.

---

## 👨‍💻 About the Author

<div align="center">

### Harsh Kushwah

Aspiring Data Analyst | Excel | SQL | Power BI

[![GitHub](https://img.shields.io/badge/GitHub-Profile-181717?style=for-the-badge&logo=github)](https://github.com/harshuu090)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=for-the-badge&logo=linkedin)](https://www.linkedin.com/in/harsh-kushwah-9346902a4/)

📧 **Email:** harshuu9752@gmail.com

</div>

---

<div align="center">

**⭐ If you find this project useful, consider giving the repository a star!**

*Built with Excel, VBA, and a focus on data-driven decision-making.*

</div>
