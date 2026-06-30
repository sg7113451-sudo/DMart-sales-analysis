:

🛒 DMart Sales Analysis
📌 Project Overview
This repository contains the DMart sales dataset and my analysis work.
The goal is to understand total sales, average net revenue, and estimate an expected marketing budget using Excel formulas.

📊 Key Learnings
Total Sales Calculation

Formula:

excel
=SUM(SalesColumn)
This adds up all sales values to get the overall revenue.

Average Net Revenue

Formula:

excel
=AVERAGE(NetRevenueColumn)
This gives the mean revenue per transaction or product.

Expected Marketing Budget

Assumption: Marketing budget is set as a percentage of total sales.

Example Formula (10% of sales):

excel
=SUM(SalesColumn)*10%
You can adjust the percentage based on business strategy.

📂 Contents
DMart_sales_dataset.csv → Raw dataset

Analysis.xlsx → Excel file with formulas applied

README.md → Documentation of methods and insights

🎨 Conditional Formatting on Discounts
To make the Discount column more insightful, I applied conditional formatting in Excel:

Highlight High Discounts (Red)

Rule: Format cells with values greater than or equal to a chosen threshold (e.g., ≥ 30%).

Formatting: Fill color Red to indicate high discounts.

Highlight Low Discounts (Green)

Rule: Format cells with values less than or equal to a chosen threshold (e.g., ≤ 5%).

Formatting: Fill color Green to indicate low discounts.

Gradient for Increasing Discounts

Used a Color Scale (Green → Yellow → Red).

As discount values increase, the cell color gradually shifts from green (low) to red (high).

Highlight Top/Bottom Discounts

Rule:

Top 10% → Highlight with bold red fill.

Bottom 10% → Highlight with bold green fill.

This makes it easy to spot extreme values at a glance.

📊 Example Formula Context
Total Sales: =SUM(SalesColumn)

Average Net Revenue: =AVERAGE(NetRevenueColumn)

Expected Marketing Budget (10%): =SUM(SalesColumn)*10%

✅ Why This Helps
Quickly identify products with unusually high discounts.

Spot low-discount items that may need promotional focus.

Visualize discount distribution across the dataset.

)

🧮 Logical Functions (IF, IFS, AND, OR, NOT, IFERROR)
Basic IF for Sales Classification

excel
=IF([@Sales]>10000,"High","Low")
Multiple Conditions with IFS

excel
=IFS([@Sales]>10000,"High",[@Sales]>5000,"Medium",[@Sales]<5000,"Low")
Combining Conditions with AND

excel
=IF(AND([@Sales]>5000,[@Profit]>500),"Good","Not Good")
Profit-to-Sales Ratio

excel
=[@Profit]/[@Sales]
Example result: 0.214289894 (≈ 21.4% margin)

Error Handling with IFERROR

excel
=IFERROR([@Profit]/[@Sales],0)
→ Returns 0 if division error occurs.


🚀 Insights
Sales above ₹10,000 are classified as High Impact.

Sales between ₹5,000–₹10,000 are Medium Impact.

Sales below ₹5,000 are Low Impact.

Profit margin helps identify Good vs Not Good performance.
