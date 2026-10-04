# 🧹 Product Dataset – Data Cleaning & Preparation Using Excel

## 📌 Project Overview

This project demonstrates practical **data cleaning and preprocessing techniques using Microsoft Excel**.

As part of my journey toward becoming a **Data Analyst**, I worked with a Product Dataset containing information such as Product ID, Product Name, Brand Name, Quantity, Category, and Price.

Real-world datasets often contain missing values, inconsistent text, duplicate records, formatting issues, and poorly structured fields. Before performing analysis, these data-quality issues need to be identified and resolved.

The objective of this project is to transform the raw dataset into a **clean, standardized, and analysis-ready dataset** using Excel.

---

## 🎯 Project Objectives

The main objectives of this project were:

* Handle missing values
* Identify and correct inconsistent text formats
* Fix category spelling errors
* Remove duplicate records
* Split and restructure columns
* Merge related columns
* Apply appropriate number and date formatting
* Use conditional formatting to improve data readability
* Prepare the dataset for further analysis

---

## 📊 Dataset Description

**Dataset Name:** Product Dataset

| Column       | Description                                              |
| ------------ | -------------------------------------------------------- |
| Product ID   | Unique identifier containing product-related information |
| Product Name | Name of the product                                      |
| Brand Name   | Brand associated with the product                        |
| Quantity     | Available quantity of the product                        |
| Category     | Product category                                         |
| Price        | Price of the product                                     |

---

# 🛠️ Data Cleaning Tasks

## 1. Handling Missing Values

### Missing Values in Price

First, the `Price` column was checked for missing values using Excel filtering and blank-cell identification.

Products with missing prices should not be assigned an arbitrary value because doing so could introduce incorrect information into the dataset.

### Strategy

Depending on the business requirement, missing prices can be handled by:

* Checking the original source for the correct price
* Using the product's current or standard price if available
* Using the median price for similar products only when appropriate
* Keeping the value blank if the correct price cannot be determined

For this project, the preferred approach is to **investigate the source data before imputing a price**.

### Missing Categories

Missing values in the `Category` column can be handled by:

* Checking the Product Name to determine the appropriate category
* Using Brand/Product information to infer the category
* Checking the original source data
* Assigning `"Unknown"` when the category cannot be reliably determined

This avoids introducing incorrect assumptions into the dataset.

---

# 2. Correcting Inconsistent Data

## Product Name Standardization

The `Product Name` column was checked for inconsistent text formats such as:

* Different capitalization
* Unnecessary spaces
* Inconsistent naming patterns
* Extra characters

Excel's **Find and Replace** functionality can be used to standardize recurring inconsistencies.

Example:

| Before               | After                |
| -------------------- | -------------------- |
| `iphone 15`          | `iPhone 15`          |
| `Samsung Galaxy s23` | `Samsung Galaxy S23` |
| `  Laptop`           | `Laptop`             |

Additional Excel functions such as `TRIM()` can be used to remove unnecessary spaces.

---

## Category Correction

The `Category` column was checked for spelling mistakes and inconsistent values.

Examples of possible corrections:

| Incorrect     | Correct       |
| ------------- | ------------- |
| `Electornics` | `Electronics` |
| `Electroncs`  | `Electronics` |
| `Clothng`     | `Clothing`    |

Excel's **Find and Replace** feature was used to correct known spelling mistakes and standardize category names.

This ensures that the same category is represented consistently throughout the dataset.

---

# 3. Removing Duplicate Records

Duplicate rows were identified by comparing the **entire row**, including:

* Product ID
* Product Name
* Brand Name
* Quantity
* Category
* Price

Excel's:

**Data → Remove Duplicates**

feature can be used to remove exact duplicate records.

Only complete duplicate rows were removed so that legitimate products with similar attributes were not accidentally deleted.

---

# 4. Splitting and Merging Data

## Splitting Product ID

The `Product ID` column contains information that needs to be separated into:

* `Manufacturing Date`
* `Country Code`

Excel's **Text to Columns** feature can be used when the Product ID follows a consistent delimiter-based structure.

Unnecessary characters were removed during the transformation.

### Result

The original:

`Product ID`

was transformed into:

| Manufacturing Date | Country Code |
| ------------------ | ------------ |
| Date value         | Country code |

This restructuring makes the information easier to analyze.

> **Note:** The exact splitting method depends on the actual structure of the Product ID in the dataset.

---

## Merging Brand Name and Product Name

The `Brand Name` and `Product Name` columns were combined into a new column called:

**Product Brand**

Example:

| Brand Name | Product Name | Product Brand        |
| ---------- | ------------ | -------------------- |
| Apple      | iPhone 15    | Apple - iPhone 15    |
| Samsung    | Galaxy S23   | Samsung - Galaxy S23 |

An Excel formula such as:

```excel
=A2&" - "&B2
```

can be used to combine the two fields.

Alternatively:

```excel
=CONCAT(A2," - ",B2)
```

---

# 5. Number Formatting

## Price Formatting

The `Price` column was formatted as **Currency** to make monetary values easier to read and interpret.

Example:

**Before:**

```text
49999
```

**After:**

```text
₹49,999.00
```

The appropriate currency symbol can be selected based on the dataset/business context.

---

## Manufacturing Date Formatting

The `Manufacturing Date` column was formatted using:

```text
DD-MM-YYYY
```

Example:

```text
15-08-2025
```

This provides a consistent date format throughout the dataset.

---

# 6. Conditional Formatting

## Price Column

Conditional formatting was applied to the `Price` column using a **Data Bar** or **Color Scale**.

This provides a visual representation of product prices and makes it easier to identify relatively high- and low-priced products.

For example:

* Shorter data bar → Lower price
* Longer data bar → Higher price

---

## Electronics Category

A custom conditional formatting rule was created for the `Category` column to highlight products belonging to the:

```text
Electronics
```

category.

Example Excel formula:

```excel
=A2="Electronics"
```

The formatting can then be applied to highlight matching cells.

This makes Electronics products easy to identify visually.

---

# 📋 Data Cleaning Workflow

The overall workflow followed in this project was:

```text
Raw Dataset
     ↓
Check Missing Values
     ↓
Standardize Text
     ↓
Correct Category Errors
     ↓
Remove Duplicates
     ↓
Split Product ID
     ↓
Merge Brand + Product Name
     ↓
Format Price & Dates
     ↓
Apply Conditional Formatting
     ↓
Clean & Analysis-Ready Dataset
```

---

# 💻 Excel Skills Demonstrated

Through this project, I practiced the following Microsoft Excel skills:

* Data Cleaning
* Missing Value Handling
* Find & Replace
* Text Standardization
* `TRIM()`
* Text to Columns
* `CONCAT()`
* Excel formulas
* Remove Duplicates
* Currency Formatting
* Date Formatting
* Conditional Formatting
* Data Bars
* Color Scales
* Data Validation and Quality Checking

---

# 📁 Project Structure

```text
Product-Dataset-Excel-Data-Cleaning/
│
├── README.md
│
├── data/
│   └── product_dataset_raw.xlsx
│
├── cleaned_data/
│   └── product_dataset_cleaned.xlsx
│
└── screenshots/
    ├── missing_values.png
    ├── duplicate_removal.png
    ├── text_standardization.png
    ├── conditional_formatting.png
    └── final_dataset.png
```

---

# 📈 Key Learning Outcomes

This project helped me understand that **data analysis begins with data quality**.

I learned how to:

1. Identify data-quality problems before analysis.
2. Handle missing values using appropriate business logic.
3. Standardize inconsistent text data.
4. Detect and remove exact duplicate records.
5. Restructure data into useful analytical columns.
6. Apply appropriate number and date formats.
7. Use conditional formatting to improve data readability.
8. Prepare raw data for further analysis and visualization.

---

# 🚀 Future Improvements

As a next step, I plan to use the cleaned dataset for:

* Exploratory Data Analysis (EDA)
* Product category analysis
* Brand-wise analysis
* Price distribution analysis
* Quantity analysis
* Excel dashboards
* Power BI visualization
* Business insights and recommendations

---

# 👨‍💻 About Me

I am an aspiring **Data Analyst** building my portfolio through practical projects involving **Excel, SQL, Power BI, Python, and Data Visualization**.

This project is part of my learning journey to develop practical data-cleaning and analytical skills using real-world-style datasets.

---

## ⭐ Skills

**Microsoft Excel | Data Cleaning | Data Preprocessing | Data Analysis | SQL | Power BI | Python | Data Visualization**

---

## 📌 Project Status

**Completed – Excel Data Cleaning & Preparation**

More data analytics projects will be added to this portfolio as I continue developing my skills.
