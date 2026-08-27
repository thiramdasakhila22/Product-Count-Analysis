# Veda Technology – Day 15 | Task 15: Product Count Analysis

## 📌 Project Overview

This project was completed as part of the **Veda Technology Data Analytics Track – Day 15, Task 15**.

The objective was to analyze product records by category using Microsoft Excel, calculate category-wise product counts, determine the largest category, and present the findings visually.

## 🎯 Objective

* Count products by category.
* Practice `COUNTIF` and Excel formulas.
* Calculate each category's percentage contribution.
* Identify the category with the highest product count.
* Present the analysis using a chart.

## 🛠️ Tools Used

* Microsoft Excel
* COUNTIF
* INDEX & MATCH
* Basic data analysis
* Column chart

## 📊 Dataset

The analysis uses the **Superstore sales dataset**.

The dataset contains sales transaction records across different product categories, including:

* Furniture
* Office Supplies
* Technology

## 🔎 Methodology

### 1. Category-wise Product Count

The `COUNTIF` function was used to count records belonging to each category.

Example:

```excel
=COUNTIF(Sales_Data!H:H,A2)
```

### 2. Percentage Contribution

The percentage of total records for each category was calculated using:

```excel
=B2/SUM($B$2:$B$4)
```

### 3. Largest Category

The category with the highest count was identified using:

```excel
=INDEX(A2:A4,MATCH(MAX(B2:B4),B2:B4,0))
```

## 📈 Results

| Category        | Product Count | % of Total |
| --------------- | ------------: | ---------: |
| Furniture       |         2,118 |        21% |
| Office Supplies |         6,012 |        60% |
| Technology      |         1,847 |        19% |

### 🏆 Top Category

**Office Supplies** is the largest category with **6,012 records**, representing approximately **60% of the total records** analyzed.

## 💡 Key Insight

Office Supplies contributes the largest share of records among the three categories. It represents approximately 60% of the analyzed records, followed by Furniture at 21% and Technology at 19%.

## 📊 Visualization

A column chart was created in Excel to compare the product count across the three categories.

## 📁 Deliverable

* `Veda_Technology_Day15_Task15_Product_Count_Analysis.xlsx` – Excel analysis containing the product count table, percentage analysis, largest category identification, chart, and key insight.

## 🧠 Skills Demonstrated

* Excel data analysis
* COUNTIF
* INDEX & MATCH
* Percentage calculations
* Data visualization
* Business insight generation
* Analytical thinking

## 👩‍💻 Project Context

**Program:** Veda Technology – Data Analytics Track
**Day:** 15
**Task:** 15 – Product Count Analysis
