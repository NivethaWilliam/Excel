
<div align="center">

# 🧹 Excel Assignment 2: Data Cleaning & Transformation

### By Nivetha William

**Microsoft Excel | Data Cleaning | Data Transformation | Data Quality**

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Data Cleaning](https://img.shields.io/badge/Data_Cleaning-1565C0?style=for-the-badge)
![Data Transformation](https://img.shields.io/badge/Data_Transformation-6A1B9A?style=for-the-badge)

</div>

---

## 📌 Project Overview

This project focuses on cleaning and transforming a product dataset using Microsoft Excel. The objective is to identify data quality issues, handle missing values, correct inconsistent entries, remove duplicate records, and prepare a structured dataset for further analysis.

This assignment demonstrates practical Excel skills that are essential for data analytics and business intelligence.

## 🎯 Objectives

- Identify and handle missing values.
- Standardize inconsistent product names and categories.
- Detect and remove duplicate records.
- Extract and transform information from Product IDs.
- Combine product names and brand names into a single column.
- Apply appropriate number and date formatting.
- Prepare a clean dataset for further analysis.

## 📂 Dataset Description

The original dataset contains product-related information.

| Column | Description |
|---|---|
| Product ID | Unique product identifier |
| Manufacturing Date | Date extracted and converted from Product ID |
| Country Code | Country identifier extracted from Product ID |
| Product Name | Name of the product |
| Brand Name | Product brand |
| Product Brand | Combined product and brand name |
| Price ($) | Product price in dollars |
| Quantity | Quantity associated with the product |
| Category | Product category |

## 🛠️ Data Cleaning & Transformation Steps

### 1. Handling Missing Values

Identified missing values in the Price and Category columns.

- Used `ISBLANK()` to identify empty cells.
- Used `AVERAGE()` and `IF()` to impute missing product prices.
- Replaced missing categories with `Uncategorized`.

**Example formulas:**

```excel
=ISBLANK(G2)
```

```excel
=IF(ISBLANK(G2),AVERAGE($G$2:$G$32),G2)
```

```excel
=IF(ISBLANK(I2),"Uncategorized",I2)
```

### 2. Correcting Inconsistent Data

Standardized product names and corrected inconsistent category entries.

Examples:
- `laptop` → `Laptop`
- `smartphone` → `Smartphone`
- `Electroni` → `Electronics`

This improved consistency and made the dataset easier to analyze.

### 3. Removing Duplicate Records

Identified duplicate product records and removed repeated entries from the final dataset.

- Checked repeated product records.
- Removed duplicate entries.
- Retained one record for each repeated entry identified during cleaning.

### 4. Splitting and Extracting Data

Extracted date and country information from the Product ID.

**Extract the date portion:**

```excel
=LEFT(A2,6)
```

**Extract the country code:**

```excel
=RIGHT(A2,2)
```

The extracted date portion was converted into a proper Excel date value in the final dataset.

### 5. Combining Product and Brand Names

Created a new `Product Brand` column by combining product names and brand names.

```excel
=CONCATENATE(D2," ",E2)
```

**Example:**

`Laptop` + `Dell` → `Laptop Dell`

### 6. Number and Date Formatting

- Converted extracted date values into proper date format.
- Maintained consistent product price and quantity formatting.
- Improved readability and consistency across the final dataset.

### 7. Conditional Formatting

Used conditional formatting to highlight data values and support easier identification of relevant records during the cleaning process.

---

## 📊 Project Summary

| Metric | Result |
|---|---:|
| Original data records | 34 |
| Final cleaned records | 31 |
| Records removed during deduplication | 3 |
| Final dataset columns | 9 |
| Missing price values identified | 3 |
| Missing category values identified | 4 |

*The record counts are based on the original and final worksheets in the uploaded workbook. Missing-value counts refer to the original dataset.*

## 📈 Final Output

The final cleaned worksheet contains standardized and structured product information, including:

- Consistent product names and categories.
- Imputed missing price values.
- Missing categories labeled as `Uncategorized`.
- Extracted manufacturing dates and country codes.
- Combined product and brand names.
- Duplicate records removed.

The cleaned dataset is better prepared for reporting, visualization, and further analysis.

## 💡 Key Learnings

- Applied Excel formulas for missing-value detection and imputation.
- Improved data consistency through standardization.
- Identified and removed duplicate records.
- Practiced text extraction and data transformation.
- Learned the importance of data quality in analytics.
- Prepared structured data for future Power BI and SQL projects.

## 🧰 Tools & Technologies

- **Microsoft Excel**
- Excel formulas: `IF`, `ISBLANK`, `AVERAGE`, `LEFT`, `RIGHT`, `CONCATENATE`
- Data cleaning and standardization
- Duplicate detection and removal
- Date and number formatting
- Conditional formatting

## 🚀 Conclusion

This assignment strengthened my practical understanding of data cleaning and transformation using Microsoft Excel.

By addressing missing values, inconsistent entries, duplicate records, and unstructured text, I learned how to prepare raw data for reliable analysis and reporting.

This project is part of my journey toward becoming a **Data Analyst**, with a growing focus on Excel, SQL, Power BI, and business analytics.

---

<div align="center">

### ✨ Clean Data. Consistent Results. Better Insights. 🚀

**Nivetha William | Aspiring Data Analyst**

[GitHub](https://github.com/NivethaWilliam) · [LinkedIn](https://www.linkedin.com/in/nivetha-william-417723351/)

</div>
