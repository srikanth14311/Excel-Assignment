# 📊 Product Dataset Analysis Using Microsoft Excel

## 1. Project Overview

This project focuses on performing basic data exploration and analysis using **Microsoft Excel**. The objective is to analyze a product dataset using Excel formulas and functions to calculate summary statistics, categorize products based on price, perform conditional calculations, and extract useful information from Product IDs.

This project demonstrates fundamental Excel skills that are essential for a beginner Data Analyst.

## 2. Dataset Description

**Dataset Name:** Product Dataset

The dataset contains information about different products, including their identification details, brand information, quantity, category, and price.

### Dataset Columns

| Column Name  | Description                                  |
| ------------ | -------------------------------------------- |
| Product ID   | Unique identification number of each product |
| Product Name | Name of the product                          |
| Brand Name   | Brand associated with the product            |
| Quantity     | Available quantity of the product            |
| Category     | Product category                             |
| Price        | Price of the product                         |

## 3. Project Objectives

The main objectives of this project are:

* Perform basic data exploration.
* Calculate total, count, and average values.
* Identify minimum and maximum product prices.
* Categorize products using logical functions.
* Perform conditional calculations.
* Extract specific information from Product IDs.
* Improve practical Excel formula skills.

## 4. Excel Functions Used

This project uses the following Excel functions:

| Function       | Purpose                                            |
| -------------- | -------------------------------------------------- |
| SUM            | Calculates the total price of all products         |
| COUNT / COUNTA | Counts numeric values or non-empty product records |
| AVERAGE        | Calculates the average product price               |
| MIN            | Identifies the minimum product price               |
| MAX            | Identifies the maximum product price               |
| IF             | Categorizes products based on price                |
| SUMIF          | Calculates the total price of Electronics products |
| COUNTIF        | Counts products with a price below $100            |
| LEFT           | Extracts the first two characters of Product ID    |
| RIGHT          | Extracts the last two characters of Product ID     |
| MID            | Extracts characters from the middle of Product ID  |

## 5. Tasks Performed

### Task 1: Basic Data Exploration

Performed the following calculations:

* Total Price: Calculated the sum of all product prices.
* Total Products: Counted the number of products in the dataset.
* Average Price: Calculated the average price of all products.

### Task 2: Minimum and Maximum Values

Identified:

* Minimum Price: Lowest product price.
* Maximum Price: Highest product price.

### Task 3: Logical Function – IF

Created a new column named **Price Range**.

Products were categorized based on the following conditions:

| Condition     | Classification |
| ------------- | -------------- |
| Price >= $500 | High Price     |
| Price < $500  | Standard Price |

**Excel Formula:**

```excel
=IF(F2>=500,"High Price","Standard Price")
```

### Task 4: Conditional Functions – SUMIF and COUNTIF

Performed conditional calculations:

**1. Total Price of Electronics Products**

Used SUMIF to calculate the total price of products belonging to the Electronics category.

```excel
=SUMIF(E2:E100,"Electronics",F2:F100)
```

**2. Count Products Below $100**

Used COUNTIF to count the number of products whose price is less than $100.

```excel
=COUNTIF(F2:F100,"<100")
```

### Task 5: Text Formatting – LEFT, RIGHT, MID

Extracted specific information from the Product ID column and created three new columns.

**1. Day – LEFT Function**

Extracted the first two characters from Product ID.

```excel
=LEFT(A2,2)
```

**2. Country Code – RIGHT Function**

Extracted the last two characters from Product ID.

```excel
=RIGHT(A2,2)
```

**3. Month – MID Function**

Extracted characters from the 4th to 6th positions of Product ID.

```excel
=MID(A2,4,3)
```

## 6. Skills Demonstrated

Through this project, I demonstrated the following skills:

* Microsoft Excel
* Data Exploration
* Data Summarization
* Logical Functions
* Conditional Aggregation
* Text Manipulation
* Basic Data Analysis

## 7. Tools Used

* **Microsoft Excel:** Data exploration, calculations, and analysis.
* **GitHub:** Project documentation and portfolio management.

## 8. Key Learning Outcomes

* Learned how to summarize datasets using Excel formulas.
* Understood how to apply logical conditions using the IF function.
* Practiced conditional calculations using SUMIF and COUNTIF.
* Learned how to extract information from text fields using LEFT, RIGHT, and MID.
* Developed foundational skills for real-world data analysis tasks.

## 9. Conclusion

This project helped develop a strong foundation in Microsoft Excel for data analysis. By applying basic mathematical, logical, conditional, and text functions, useful information was extracted and summarized from the product dataset.

These skills provide a foundation for more advanced data analysis techniques using Excel, SQL, Power BI, and Python.

---

**Project Type:** Beginner-Level Data Analysis Project
**Domain:** Product Data Analysis
**Tool:** Microsoft Excel
**Author:** Data Analyst Aspirant

