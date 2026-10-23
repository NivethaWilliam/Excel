
<div align="center">

# 📊 Excel Assignment 1: Data Exploration

### By Nivetha William

**Microsoft Excel | Data Analysis | Excel Formulas | Data Cleaning**

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data_Analysis-1565C0?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)

</div>

---

## 📌 Project Overview

This project focuses on exploring and analyzing a product dataset using Microsoft Excel. The assignment demonstrates essential Excel functions for numerical analysis, conditional categorization, and text extraction.

The objective is to develop practical Excel skills and understand how spreadsheet formulas can transform raw data into meaningful business insights.

## 🎯 Objectives

- Calculate summary statistics using Excel formulas.
- Identify minimum and maximum product prices.
- Categorize products based on their price.
- Analyze product prices using conditional functions.
- Extract useful information from Product IDs using text functions.

## 📂 Dataset Description

The dataset contains product information, including:

| Column | Description |
|---|---|
| Day | Day extracted from Product ID |
| Country Code | Country identifier extracted from Product ID |
| Month | Month abbreviation extracted from Product ID |
| Product ID | Unique product identifier |
| Product Name | Name of the product |
| Brand Name | Brand associated with the product |
| Price ($) | Product price in dollars |
| Quantity | Product quantity |
| Category | Product category |
| Price Range | Product price classification |

## 🛠️ Excel Functions Used

### 1. SUM, COUNT, and AVERAGE

Used to calculate the total price, number of products, and average product price.

```excel
=SUM(G2:G35)
=COUNT(G2:G35)
=AVERAGE(G2:G35)
```

### 2. MIN and MAX

Used to identify the lowest and highest product prices.

```excel
=MIN(G2:G35)
=MAX(G2:G35)
```

### 3. IF Function

Created a `Price Range` column to categorize products priced at $500 or above as High Price and the remaining products as Standard Price.

```excel
=IF(G2>=500,"High Price","Standard Price")
```

### 4. SUMIF and COUNTIF

Calculated the total price of products in the Electronics category and counted products priced below $100.

```excel
=SUMIF(I2:I35,"Electronics",G2:G35)
=COUNTIF(G2:G35,"<100")
```

### 5. Text Functions: LEFT, RIGHT, and MID

Extracted the day, country code, and month from the Product ID.

**Extract Day:**

```excel
=LEFT(D2,2)
```

**Extract Country Code:**

```excel
=RIGHT(D2,2)
```

**Extract Month:**

```excel
=MID(D2,4,3)
```

## 📈 Key Results

Based on the uploaded dataset:

| Analysis | Result |
|---|---:|
| Total product price | $10,100 |
| Number of products | 34 |
| Average product price | $297.06 |
| Minimum product price | $30 |
| Maximum product price | $1,000 |
| Total price of Electronics products | $8,050 |
| Products priced below $100 | 11 |
| Products classified as High Price | 8 |

*Note: The total product price is the sum of listed product prices, not the total inventory value calculated using quantity.*

## 💡 Key Learnings

- Applied Excel formulas to perform numerical analysis.
- Used conditional logic to classify products.
- Analyzed product data using SUMIF and COUNTIF.
- Extracted structured information from text using LEFT, RIGHT, and MID.
- Improved data exploration and spreadsheet problem-solving skills.

## 🧰 Tools & Technologies

- **Microsoft Excel**
- Excel formulas and functions
- Data exploration and analysis
- Product data categorization
- Text manipulation

## 🚀 Conclusion

This assignment strengthened my understanding of Microsoft Excel and its application in data analysis. By using mathematical, conditional, and text functions, I transformed product data into structured information and generated useful summary insights.

This project is part of my journey toward becoming a **Data Analyst**, building practical skills in Excel, SQL, Power BI, and business analytics.

---

<div align="center">

**Learning Excel. Exploring Data. Building Analytical Skills. 🚀**

</div>
