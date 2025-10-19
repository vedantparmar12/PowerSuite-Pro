---
name: Excel Master Controller
description: Create and manipulate Excel spreadsheets with formulas, charts, and formatting. Use when user needs spreadsheet creation, data analysis, budgets, or Excel automation.
version: 1.0.0
dependencies: python>=3.8, openpyxl>=3.1.0, pandas>=1.5.0, xlsxwriter>=3.0.0, pillow>=9.0.0
---

# Excel Master Controller

Comprehensive Excel automation skill providing complete spreadsheet control from single prompts. Creates professional spreadsheets with intelligent data organization, advanced formulas, dynamic charts, and automated formatting.

## Quick Start

```python
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.chart import BarChart, LineChart, PieChart
import pandas as pd

# Create workbook with intelligent structure
workbook = openpyxl.Workbook()
# Full implementation in scripts/excel_master.py
```

## Core Capabilities

### 1. Intelligent Data Organization
- **Schema Detection**: Automatically identify data types and relationships
- **Smart Structuring**: Organize data with appropriate headers, sections, and hierarchies
- **Dynamic Layouts**: Adapt table structure based on content complexity
- **Data Validation**: Built-in validation rules for data integrity

### 2. Advanced Formula Engine  
- **Formula Generation**: Create complex formulas from natural language descriptions
- **Cross-Sheet References**: Link data across multiple worksheets intelligently
- **Dynamic Calculations**: Formulas that adapt as data changes
- **Business Logic**: Implement conditional logic, lookups, and aggregations

### 3. Professional Formatting System
- **Conditional Formatting**: Highlight patterns, trends, and exceptions automatically
- **Corporate Styling**: Apply consistent brand colors, fonts, and layouts
- **Data Visualization**: Format cells for optimal readability and impact
- **Print Optimization**: Ensure professional appearance when printed

### 4. Chart and Visualization Engine
- **Smart Chart Selection**: Choose optimal chart types based on data characteristics
- **Interactive Dashboards**: Create linked charts that update dynamically
- **Custom Visualizations**: Specialized charts for specific business needs
- **Data Storytelling**: Arrange visuals to tell compelling data stories

### 5. On-Demand Control Features
- **Real-Time Updates**: Modify existing sheets without recreating from scratch  
- **Selective Operations**: Target specific ranges, sheets, or data elements
- **Batch Processing**: Apply changes across multiple files or sheets
- **Version Control**: Track changes and maintain audit trails

## Usage Patterns

### Complete Spreadsheet Creation
```
User: "Create a quarterly sales tracking sheet with team performance metrics"
→ Generates: Structured data entry, calculated KPIs, trend charts, 
  conditional formatting, summary dashboard
```

### On-Demand Updates
```
User: "Add a new product category column and update all formulas"
→ Modifies: Existing structure, recalculates dependencies, 
  updates charts, maintains formatting consistency
```

### Data Analysis Automation
```
User: "Analyze this data and create executive summary with key insights"
→ Produces: Statistical analysis, trend identification, 
  executive dashboard, automated recommendations
```

## Specialized Functions

### Financial Management
- **Budget Templates**: Comprehensive budget tracking with variance analysis
- **P&L Statements**: Automated profit & loss calculations with drill-downs
- **Cash Flow**: Dynamic cash flow projections with scenario modeling
- **ROI Analysis**: Investment return calculations with sensitivity analysis

### Project Management  
- **Gantt Charts**: Visual project timelines with dependency tracking
- **Resource Planning**: Capacity planning and allocation optimization
- **Progress Tracking**: Milestone monitoring with automated status updates
- **Risk Registers**: Risk assessment matrices with mitigation tracking

### Data Analytics
- **Pivot Table Automation**: Intelligent data summarization and grouping
- **Statistical Analysis**: Regression, correlation, and trend analysis
- **Forecasting Models**: Predictive analytics with confidence intervals
- **KPI Dashboards**: Executive dashboards with real-time metrics

## File Organization

- `SKILL.md` - Main instructions (this file)
- `FORMULAS.md` - Advanced formula library and templates
- `FORMATTING.md` - Professional styling guidelines and templates
- `CHARTS.md` - Chart creation and customization guide
- `scripts/excel_master.py` - Core Excel automation engine
- `scripts/formula_builder.py` - Dynamic formula generation system
- `scripts/chart_creator.py` - Intelligent chart and visualization engine
- `scripts/data_analyzer.py` - Advanced data analysis and insights
- `templates/` - Pre-built Excel templates for common use cases
- `samples/` - Example datasets and completed spreadsheets

## Advanced Features

### Intelligent Data Processing
- **Auto-Detection**: Identify data patterns, outliers, and relationships
- **Smart Suggestions**: Recommend formulas, charts, and formatting
- **Error Prevention**: Validate data integrity and formula consistency  
- **Performance Optimization**: Efficient calculations for large datasets

### Business Intelligence Integration
- **External Data**: Connect to databases, APIs, and external sources
- **Automated Reporting**: Schedule regular report generation and distribution
- **Multi-Source Analysis**: Combine data from multiple files and sources
- **Export Flexibility**: Generate reports in multiple formats (PDF, CSV, etc.)

### Collaboration Features
- **Multi-User Setup**: Design sheets for team collaboration
- **Access Control**: Protect sensitive data while enabling sharing
- **Change Tracking**: Monitor and log all modifications
- **Comments Integration**: Embed contextual notes and explanations

## Integration Capabilities

- Works seamlessly with PowerPoint skill for data-driven presentations
- Imports data from various sources (CSV, JSON, databases)
- Exports results in multiple formats for further processing
- Supports automated email distribution of reports
- Integrates with cloud storage platforms

## Usage Examples

### Simple Creation
```
"Create expense tracking spreadsheet"
→ Produces categorized expense tracker with totals and monthly summaries
```

### Advanced Analytics  
```
"Analyze sales data for trends and create executive dashboard"
→ Generates comprehensive analysis with predictive insights and visualizations
```

### On-Demand Modifications
```
"Update the Q3 budget sheet with new department allocations"
→ Modifies existing structure while preserving formulas and formatting
```

For advanced formulas and functions, see [FORMULAS.md](FORMULAS.md)
For professional styling guidelines, see [FORMATTING.md](FORMATTING.md)
For chart creation best practices, see [CHARTS.md](CHARTS.md)