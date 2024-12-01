# Financial Statement Analysis with Power Pivot

This repository contains a comprehensive financial analysis project designed to leverage Power Pivot and advanced DAX calculations for insightful decision-making. Below, we outline the project's structure, methodology, and tools used.

## Project Overview
The project is centered around analyzing financial performance using:

1. **Income Statement (Profit & Loss) Metrics**  
   Provided detailed DAX measures for revenue, gross margin, operating profit, and net income.
   ![Capture 03](https://github.com/user-attachments/assets/2fa1ea9f-cfc2-45ac-8f2b-e3d72e2e8f77)


3. **Balance Sheet Analysis**  
   Included assets, liabilities, and equity breakdown to track financial health.
   ![Capture 02](https://github.com/user-attachments/assets/65184bc2-d266-43ac-840b-136478eb8155)


5. **Custom Ratios**  
   Calculated liquidity, profitability, and efficiency metrics to assess company performance.
   ![Capture 01](https://github.com/user-attachments/assets/8d87ba39-9150-4347-a649-c5d72002f24b)


## How Data Was Handled

### Data Preparation
1. **Source Data Compilation**: Financial data was gathered and organized into structured tables in Excel. 
2. **Cleaning and Transformation**: Used Power Query to:
   - Remove duplicates.
   - Standardize date formats and column names.
   - Handle missing values effectively.

### Building the Data Model with Power Pivot
1. **Data Loading**: Imported cleaned tables into Power Pivot to maintain relationships.
2. **Relationships Setup**: Established connections between fact and dimension tables, including:
   - Fact Table: Transactions.
   - Dimension Tables: Dates, Accounts, and Departments.
  
   - ![Capture 04](https://github.com/user-attachments/assets/d04ea276-82bc-4ea8-9397-994fc9301f93)

3. **DAX Calculations**:
   - Created calculated columns for KPIs (e.g., revenue growth).
   - Built measures to aggregate data dynamically (e.g., Total Sales, Profit Margin).
  
   - # Financial Statement Analysis - DAX Calculation Template

## Overview
This template provides a comprehensive set of DAX calculations for financial statement analysis, covering Income Statement, Balance Sheet, Cash Flow, and Key Financial Ratios.

## 1. Income Statement (Profit & Loss) Metrics

### Revenue Metrics
```dax
// Total Sales Calculations
Total_Sales = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Sales")
Sales_YTD = TOTALYTD([Total_Sales], Bi_Calendar[Date])
```

### Profitability Metrics
```dax
// Profit Calculations
Gross_Profit = CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Gross Profit")
Gross_Profit_Margin = [Gross_Profit] / [Total_Sales]

EBITDA = [Gross_Profit] + CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Operating Expenses")
Operating_Profit = CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Operating Profit")
Operating_Profit_Margin = [EBITDA] / CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Depreciation & Amortization")

Net_Profit = CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Net Profit")
Net_Profit_Margin = [Net_Profit] / [Total_Sales]
```

### Expense Metrics
```dax
// Expense Calculations
Cost_of_Sales = CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Cost of Sales")
Marketing_Cost = CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Marketing")
Interest_Expensed = CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Interest Expense")
```

## 2. Balance Sheet Metrics

### Asset Metrics
```dax
// Asset Calculations
Total_Assets = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Assets")
Current_Assets = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Current Assets")
Inventory = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Inventory")
Receivables = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Trade Receivables")
```

### Liability Metrics
```dax
// Liability Calculations
Total_Liabilities = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Liabilities")
Current_Liabilities = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Current Liabilities")
Payables = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Trade Payables")
```

### Equity Metrics
```dax
// Equity Calculations
Total_Equity = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Owners Equity")
Capital_Employed = CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass] = "Owners Equity") + 
                   CALCULATE([Value_Total_to_Date (TTD)], 'bl_ChartofAccounts'[SubClass2] = "Long Term Liabilities")
```

## 3. Financial Ratios and Performance Metrics

### Liquidity Ratios
```dax
// Liquidity Calculations
Current_Ratio = [Current_Assets] / [Current_Liabilities]
Quick_Ratio = ([Current_Assets] - [Inventory]) / [Current_Liabilities]
```

### Efficiency Ratios
```dax
// Efficiency Calculations
Asset_Turnover = [Total_Sales] / [Total_Assets]
Inventory_Turnover_Period = [Inventory] / [Cost_of_Sales] * 365 - 1
Receivables_Turnover_Period = (CALCULATE([Total_Sales], 'bl_ChartofAccounts'[SubClass] = "Sales") * 365) / [Receivables]
Payables_Payment_Period = (CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Cost of Sales") * 365) - 1
```

### Profitability Ratios
```dax
// Profitability Calculations
Return_on_Equity_ROE = [Net_Profit] / [Total_Equity]
Return_on_Capital_Employed_ROCE = [PBIT] / ([Capital_Employed] + [Current_Liabilities])
Gross_Profit_Margin = [Gross_Profit] / [Total_Sales]
Net_Profit_Margin = [Net_Profit] / [Total_Sales]
```

### Leverage Ratios
```dax
// Leverage Calculations
Gearing = CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Trading account") / [Total_Equity]
Interest_Cover = [EBIT] / [Interest_Expensed]
PBIT = [Operating_Profit] - CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "Non-operating")
```

## 4. Cash Flow and Time-Based Metrics

### Cash Flow Metrics
```dax
// Cash Flow Calculations
Cash_Flow_Statement_Value = 
    CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass] = "FTP") 
    + CALCULATE([Total_Value], 'bl_ChartofAccounts'[SubClass2] = "FTP_Negative_CF_S") 
    - CALCULATE([Total_FTP_Positive_CF], 'bl_ChartofAccounts'[SubClass] = "FTP_Positive")
```

### Moving Averages and Growth Metrics
```dax
// Time-Based Calculations
Month_Over_Month_Growth_Rate = DIVIDE([Total_Value], [Previous Month Value FTP], "-") - 1
Previous_Month_Value_FTP = CALCULATE([Total_Value], DATEADD(Bi_Calendar[Date], -1, MONTH)) / [Total_Value]

30_Day_Rolling_Value = CALCULATE(Total_Value, DATESEETWEEN(Bi_Calendar[Date], [Max Date] - 90, [Max Date]))
30_Days_Moving_Average = [90_Day_Rolling_Value] / 30

03_Month_Rolling = CALCULATE(Total_Value, DATESINPERIOD(Date, [Calendar[Date]], MAX(Bi_Calendar[Date]), -3, MONTH))
03_Month_Moving_Average = CALCULATE([03_Month_Rolling] / 3, DATESINPERIOD(Date, [Calendar[Date]], MAX(Bi_Calendar[Date]), -3, MONTH))

// Date-Related Helper Measures
Max_Date = MAX(Bi_Calendar[Date])
Min_Date = MIN(Bi_Calendar[Date])
Number_of_Months = DISTINCT(COUNT(Bi_Calendar[Date (Month)]))
```

These calculations are provided as a template and should be carefully validated and customized to your specific business context.

4. **Hierarchies and Filters**: Defined hierarchies for drill-down analysis and applied slicers for dynamic filtering.

## Key Highlights of Power Pivot Usage
- **Optimized Performance**: Leveraged Power Pivot's in-memory processing to handle large datasets efficiently.
- **Dynamic Insights**: Designed interactive dashboards that update instantly with changes in input data.
- **Scalability**: Ensured the model could integrate additional metrics or data sources seamlessly.

## Insights Generated
The model provided:
- **Trend Analysis**: Monthly and yearly trends in income and expenses.
- **Variance Analysis**: Comparison of actuals vs. budgets.
- **Ratio Analysis**: Key financial ratios to evaluate liquidity, profitability, and solvency.

---

This project showcases the power of combining Excel's advanced tools like Power Pivot and DAX for robust financial analysis. Feel free to explore and adapt the template to your financial reporting needs!

