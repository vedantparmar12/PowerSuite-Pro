# Enhanced Excel Master - Complete Customization Guide

## Overview

The Enhanced Excel Master provides **full manual control** over every aspect of your spreadsheets. With 6+ professional themes, advanced features including pivot tables, conditional formatting, data validation, charts, and editing capabilities, you can create powerful Excel workbooks tailored to your exact needs.

## Table of Contents

1. [Getting Started](#getting-started)
2. [Available Themes](#available-themes)
3. [Sheet Types](#sheet-types)
4. [Configuration Structure](#configuration-structure)
5. [Editing Existing Workbooks](#editing-existing-workbooks)
6. [Advanced Features](#advanced-features)
7. [Charts and Visualization](#charts-and-visualization)
8. [Conditional Formatting](#conditional-formatting)
9. [Data Validation](#data-validation)
10. [Examples](#examples)

---

## Getting Started

### Installation

```bash
# Ensure required packages are installed
pip install openpyxl pandas

# Run the enhanced Excel master
python excel_master_enhanced.py config.json
```

### Basic Configuration

Create a JSON configuration file with your workbook specifications:

```json
{
  "theme": "corporate_blue",
  "output_path": "my_workbook.xlsx",
  "sheets": [
    {
      "name": "Data",
      "type": "data",
      "headers": ["Name", "Value", "Status"],
      "data": [
        ["Item 1", 100, "Active"],
        ["Item 2", 200, "Pending"]
      ]
    }
  ]
}
```

---

## Available Themes

### 1. Corporate Blue
**Best for:** Business reports, financial statements, formal documents
- **Primary:** Deep blue (#0033CC)
- **Secondary:** Light blue (#99CCFF)
- **Accent:** Orange (#FF6600)
- **Success:** Green (#00CC00)
- **Warning:** Yellow (#FF9900)
- **Danger:** Red (#CC0000)

### 2. Modern Dark
**Best for:** Tech companies, modern dashboards, contemporary designs
- **Primary:** Charcoal (#2D3436)
- **Secondary:** Gray (#636E72)
- **Accent:** Light blue (#74B9FF)
- **Success:** Teal (#00B894)

### 3. Financial Green
**Best for:** Financial services, banking, investment reports
- **Primary:** Dark green (#155724)
- **Secondary:** Green (#28A745)
- **Accent:** Light green (#85D5A8)

### 4. Tech Purple
**Best for:** Tech startups, SaaS companies, creative tech
- **Primary:** Purple (#6C5CE7)
- **Secondary:** Light purple (#A29BFE)
- **Accent:** Pink (#FD79A8)

### 5. Elegant Monochrome
**Best for:** Professional services, legal, conservative industries
- **Primary:** Black (#000000)
- **Secondary:** Gray (#6C757D)
- **Accent:** Light gray (#ADB5BD)

### 6. Ocean Breeze
**Best for:** Marine industries, environmental reports, clean designs
- **Primary:** Ocean blue (#0077B6)
- **Secondary:** Sky blue (#00B4D8)
- **Accent:** Light blue (#90E0EF)

---

## Sheet Types

### 1. Data Sheet
Standard data table with headers, formatting, and optional charts.

```json
{
  "name": "Sales_Data",
  "type": "data",
  "headers": ["Date", "Product", "Amount"],
  "data": [
    ["2024-01-01", "Product A", 1000],
    ["2024-01-02", "Product B", 1500]
  ],
  "formats": {
    "Amount": "$#,##0"
  }
}
```

### 2. Dashboard Sheet
Interactive dashboard with KPIs and charts.

```json
{
  "name": "Dashboard",
  "type": "dashboard",
  "title": "Executive Dashboard",
  "kpis": [
    {
      "name": "Total Revenue",
      "value": 1000000,
      "format": "$#,##0",
      "change": 0.15
    }
  ],
  "charts": [...]
}
```

### 3. Pivot Sheet
Sheet configured for pivot table-like analysis.

```json
{
  "name": "Analysis",
  "type": "pivot",
  "source_sheet": "Data",
  "pivot_config": {...}
}
```

### 4. Chart Sheet
Dedicated sheet for visualizations.

```json
{
  "name": "Charts",
  "type": "chart",
  "charts": [...]
}
```

---

## Configuration Structure

### Complete Configuration Options

```json
{
  "theme": "corporate_blue",
  "output_path": "output.xlsx",
  "sheets": [
    {
      "name": "Sheet1",
      "type": "data",
      "headers": ["Col1", "Col2", "Col3"],
      "data": [[...], [...]],
      "formats": {
        "Col2": "$#,##0",
        "Col3": "0.00%"
      },
      "formulas": [
        {
          "cell": "D2",
          "formula": "=B2+C2",
          "format": "#,##0"
        }
      ],
      "conditional_formatting": [...],
      "charts": [...],
      "validations": [...],
      "styling": {
        "table_style": true,
        "auto_width": true
      }
    }
  ]
}
```

### Custom Theme

Define a custom theme:

```json
{
  "theme": {
    "primary": "FF0033CC",
    "secondary": "FF99CCFF",
    "accent": "FFFF6600",
    "success": "FF00CC00",
    "warning": "FFFF9900",
    "danger": "FFCC0000",
    "text": "FF333333",
    "background": "FFF8F9FA",
    "header_font": "Calibri",
    "body_font": "Arial",
    "header_size": 11,
    "body_size": 10
  }
}
```

---

## Editing Existing Workbooks

The enhanced Excel master can modify existing files!

### Edit Configuration

```json
{
  "edit_file": "existing.xlsx",
  "change_theme": "modern_dark",
  "update_sheets": {
    "Sheet1": {
      "cells": {
        "A1": "Updated Title",
        "B5": 1000
      },
      "range": {
        "start": "D2",
        "data": [[1, 2], [3, 4]]
      },
      "clear": ["E10", "F10"]
    }
  },
  "add_sheets": [...],
  "delete_sheets": ["OldSheet"],
  "add_charts": {...},
  "output_path": "edited.xlsx"
}
```

### Edit Operations

#### Update Cells
Modify individual cells:
```json
"update_sheets": {
  "Sheet1": {
    "cells": {
      "A1": "New Value",
      "B2": 500,
      "C3": "=A1+B2"
    }
  }
}
```

#### Update Ranges
Modify entire ranges:
```json
"update_sheets": {
  "Sheet1": {
    "range": {
      "start": "A2",
      "data": [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
      ]
    }
  }
}
```

#### Clear Cells
Remove content:
```json
"update_sheets": {
  "Sheet1": {
    "clear": ["A1", "B2", "C3"]
  }
}
```

---

## Advanced Features

### Formulas

Add Excel formulas to cells:

```json
"formulas": [
  {
    "cell": "D2",
    "formula": "=SUM(B2:C2)",
    "format": "$#,##0"
  },
  {
    "cell": "E2",
    "formula": "=D2*0.1",
    "format": "0.00%"
  },
  {
    "cell": "F2",
    "formula": "=AVERAGE(B:B)",
    "format": "#,##0.00"
  }
]
```

### Number Formats

Common number formats:

- **Currency:** `$#,##0` or `$#,##0.00`
- **Percentage:** `0%` or `0.00%`
- **Date:** `yyyy-mm-dd` or `mm/dd/yyyy`
- **Number with commas:** `#,##0` or `#,##0.00`
- **Scientific:** `0.00E+00`
- **Text:** `@`

### Auto Width Columns

Automatically adjust column widths:

```json
"styling": {
  "auto_width": true
}
```

---

## Charts and Visualization

### Chart Types

1. **Bar Chart**
```json
{
  "type": "bar",
  "title": "Sales by Product",
  "data_range": "Sheet1!B2:B10",
  "categories_range": "Sheet1!A2:A10",
  "position": "E2",
  "width": 15,
  "height": 10,
  "x_axis": "Products",
  "y_axis": "Sales ($)",
  "style": 10
}
```

2. **Line Chart**
```json
{
  "type": "line",
  "title": "Revenue Trend",
  "data_range": "Sheet1!B1:B13",
  "categories_range": "Sheet1!A2:A13"
}
```

3. **Pie Chart**
```json
{
  "type": "pie",
  "title": "Market Share",
  "data_range": "Sheet1!B2:B5",
  "categories_range": "Sheet1!A2:A5"
}
```

4. **Area Chart**
```json
{
  "type": "area",
  "title": "Cumulative Growth"
}
```

5. **Scatter Chart**
```json
{
  "type": "scatter",
  "title": "Correlation Analysis"
}
```

6. **3D Charts**
```json
{
  "type": "bar_3d",
  "title": "3D Bar Chart"
}
```

### Chart Styling

- **Style numbers:** 1-48 (different color schemes and layouts)
- **Common styles:**
  - `10` - Corporate blue
  - `11` - Orange accent
  - `12` - Green accent
  - `26` - Monochrome
  - `42` - Colorful

---

## Conditional Formatting

### Color Scales

Three-color gradient based on values:

```json
{
  "type": "color_scale",
  "range": "B2:B20",
  "start_color": "FFE17055",
  "mid_color": "FFFECA57",
  "end_color": "FF00B894"
}
```

### Data Bars

Horizontal bars in cells:

```json
{
  "type": "data_bar",
  "range": "C2:C20",
  "color": "6C5CE7"
}
```

### Icon Sets

Icons indicating performance:

```json
{
  "type": "icon_set",
  "range": "D2:D20",
  "icon_style": "3Arrows"
}
```

**Available icon styles:**
- `3Arrows` - Up, sideways, down arrows
- `3TrafficLights1` - Red, yellow, green circles
- `3Symbols` - Check, exclamation, X
- `4Arrows` - Four directions
- `5Arrows` - Five levels

### Cell Value Rules

Highlight cells based on conditions:

```json
{
  "type": "cell_value",
  "range": "E2:E20",
  "operator": "greaterThan",
  "value": 1000,
  "color": "FF00FF00"
}
```

**Available operators:**
- `greaterThan`
- `lessThan`
- `greaterThanOrEqual`
- `lessThanOrEqual`
- `equal`
- `notEqual`
- `between`
- `notBetween`

---

## Data Validation

### Dropdown Lists

Create dropdown selections:

```json
{
  "type": "list",
  "range": "B2:B100",
  "formula1": "\"Option 1,Option 2,Option 3\"",
  "allow_blank": true,
  "error_title": "Invalid Entry",
  "error_message": "Please select from the dropdown"
}
```

### Number Validation

Restrict to valid numbers:

```json
{
  "type": "number",
  "range": "C2:C100",
  "operator": "between",
  "min": 0,
  "max": 100
}
```

### Date Validation

Ensure valid dates:

```json
{
  "type": "date",
  "range": "A2:A100",
  "operator": "greaterThan",
  "formula1": "2024-01-01"
}
```

---

## Dashboards and KPIs

### Creating Dashboards

```json
{
  "name": "Dashboard",
  "type": "dashboard",
  "title": "Q4 2024 Dashboard",
  "kpis": [
    {
      "name": "Revenue",
      "value": 1250000,
      "format": "$#,##0",
      "change": 0.15
    },
    {
      "name": "Profit Margin",
      "value": 0.28,
      "format": "0.0%",
      "change": 0.05
    },
    {
      "name": "Customers",
      "value": 1847,
      "format": "#,##0",
      "change": 0.12
    }
  ],
  "charts": [...]
}
```

### KPI Cards

Each KPI card displays:
- **Name:** Metric name
- **Value:** Current value
- **Change:** % change (optional)
- **Format:** Number format

**Change indicators:**
- Positive change: Green color
- Negative change: Red color

---

## Heatmaps

Create heatmap visualization:

```python
master.create_heatmap(
    sheet=sheet,
    data_range="B2:F20",
    title="Sales Heatmap"
)
```

This applies a color scale with:
- **White:** Minimum values
- **Orange:** Middle values
- **Red:** Maximum values

---

## Examples

### Example 1: Financial Dashboard

```bash
python excel_master_enhanced.py examples/financial_dashboard_config.json
```

Creates a comprehensive financial dashboard with:
- Executive dashboard with 4 KPIs
- Revenue tracking with trends
- Expense analysis
- Cash flow statement
- Multiple charts and conditional formatting

### Example 2: Sales Tracker

```bash
python excel_master_enhanced.py examples/sales_tracker_config.json
```

Creates a sales performance tracker with:
- Detailed sales data with validation
- Regional analysis with pie chart
- Product performance metrics
- Salesperson leaderboard
- Data bars and color scales

### Example 3: Edit Existing Workbook

```bash
python excel_master_enhanced.py examples/edit_workbook_config.json
```

Modifies an existing workbook:
- Changes theme to Ocean Breeze
- Updates specific cells and ranges
- Adds new metrics sheet
- Adds charts to multiple sheets
- Removes old sheets

---

## Tips & Best Practices

### Theme Selection

- **Financial/Corporate:** Use `corporate_blue` or `financial_green`
- **Tech/Modern:** Use `tech_purple` or `modern_dark`
- **Professional:** Use `elegant_mono`
- **Creative:** Use `ocean_breeze`

### Data Organization

- **Headers in row 1:** Always use row 1 for headers
- **Data starts row 2:** Begin data in row 2
- **Consistent formats:** Apply same format to entire columns
- **Use formulas:** Leverage Excel formulas for calculations
- **Auto-width columns:** Let the system adjust widths

### Conditional Formatting

- **Color scales:** Best for showing distribution
- **Data bars:** Great for comparing values
- **Icon sets:** Perfect for status indicators
- **Cell rules:** Highlight exceptions

### Charts

- **Bar charts:** Comparing categories
- **Line charts:** Showing trends over time
- **Pie charts:** Showing proportions (max 7 slices)
- **Area charts:** Cumulative trends
- **Scatter charts:** Correlation analysis

### Dashboards

- **2-4 KPIs per row:** Don't overcrowd
- **Charts below KPIs:** Natural flow
- **Consistent colors:** Match theme
- **Clear titles:** Descriptive names
- **Update timestamps:** Show when generated

---

## Troubleshooting

### Issue: Theme not applying
**Solution:** Check theme name spelling. Use `python excel_master_enhanced.py` to see available themes.

### Issue: Chart not displaying
**Solution:** Verify data_range and categories_range use proper Excel notation (`Sheet!A1:B10`).

### Issue: Formulas not calculating
**Solution:** Ensure formulas start with `=` and use proper cell references.

### Issue: Conditional formatting not working
**Solution:** Check range notation and color codes (must be 8-digit hex with FF prefix).

### Issue: Data validation not appearing
**Solution:** Verify the range is correct and formula1 is properly formatted.

---

## Advanced Techniques

### Dynamic Ranges

Use Excel table references for dynamic ranges:

```json
"data_range": "Data!Table1[Revenue]"
```

### Named Ranges

Reference named ranges in formulas:

```json
"formula": "=SUM(SalesData)"
```

### Multi-Sheet Formulas

Reference data across sheets:

```json
"formula": "=SUM(Sheet1!B:B)+SUM(Sheet2!B:B)"
```

### Array Formulas

Use array formulas for complex calculations:

```json
"formula": "=SUMPRODUCT(B2:B10, C2:C10)"
```

---

## Integration with Other Tools

### Export to PDF

After creating Excel file, export to PDF:

```bash
# Using LibreOffice
libreoffice --headless --convert-to pdf workbook.xlsx
```

### Import from CSV

Load CSV data into configuration:

```python
import pandas as pd
df = pd.read_csv('data.csv')
data = df.values.tolist()
headers = df.columns.tolist()
```

### Connect to Databases

Query database and generate report:

```python
import sqlite3
conn = sqlite3.connect('database.db')
df = pd.read_sql_query("SELECT * FROM sales", conn)
```

---

## Next Steps

1. **Try the examples:** Run the example configs to see different styles
2. **Create your config:** Start with an example and modify it
3. **Experiment with features:** Try conditional formatting, charts, validation
4. **Build dashboards:** Create interactive executive dashboards
5. **Automate workflows:** Generate reports programmatically

---

## Quick Reference

### Common Number Formats
```
$#,##0         → $1,000
$#,##0.00      → $1,000.50
0%             → 25%
0.00%          → 25.50%
#,##0          → 1,000
0.00           → 1.50
yyyy-mm-dd     → 2024-01-15
mm/dd/yyyy     → 01/15/2024
```

### Chart Styles Quick Guide
```
10-15   → Corporate blues and greens
16-20   → Warm oranges and reds
26-30   → Monochrome grays
35-40   → Colorful mixed
42-48   → Vibrant gradients
```

### Color Hex Codes (with FF prefix)
```
FF0033CC → Blue
FF00CC00 → Green
FFFF6600 → Orange
FFCC0000 → Red
FF6C5CE7 → Purple
```

---

**Happy Spreadsheeting! 📊💚🚀**
