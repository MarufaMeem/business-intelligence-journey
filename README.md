# Business Intelligence Journey 

This repository documents my journey from Excel beginner to Business Intelligence Analyst.

The goal is to learn and apply:

- Microsoft Excel
- SQL
- Power BI
- Data Analysis
- KPI Reporting
- Dashboard Design

Each lesson contains:
- Objective
- Dataset
- Step-by-step instructions
- Expected output
- Business use case
# Project 01 - Excel Basics

## Objective

Learn how to calculate totals, averages, highest values, and lowest values using Excel.

---

## Step 1: Create the Dataset

Open Microsoft Excel.

Create a new workbook.

Enter the following data:

| Employee | Department | Salary |
|----------|------------|---------|
| Rahim | Sales | 25000 |
| Karim | HR | 30000 |
| Jannat | Sales | 40000 |
| Nila | IT | 50000 |
| Rafi | HR | 35000 |

Your sheet should look similar to:

| A | B | C |
|---|---|---|
| Employee | Department | Salary |
| Rahim | Sales | 25000 |
| Karim | HR | 30000 |
| Jannat | Sales | 40000 |
| Nila | IT | 50000 |
| Rafi | HR | 35000 |

---

## Understanding Cell References

Excel identifies cells using:

- Column letter (A, B, C, D...)
- Row number (1, 2, 3, 4...)

Examples:

| Cell | Meaning |
|--------|---------|
| A1 | Employee |
| B2 | Sales |
| C4 | 40000 |

The salary values are located in:

```text
C2
C3
C4
C5
C6
```

Therefore:

```text
C2:C6
```

means:

```text
Start at C2 and include every cell until C6.
```

---

# Exercise 1: Calculate Total Salary

### Step 1

Click on cell:

```text
C8
```

### Step 2

Type:

```excel
=SUM(C2:C6)
```

### Step 3

Press:

```text
Enter
```

### Result

Excel should display:

```text
180000
```

Explanation:

```text
25000
+30000
+40000
+50000
+35000
--------
180000
```

---

# Exercise 2: Calculate Average Salary


```excel
=AVERAGE(C2:C6)
```

### Result

```text
36000
```

---

# Exercise 3: Find Highest Salary


```excel
=MAX(C2:C6)
```

### Result

```text
50000
```

---

# Exercise 4: Find Lowest Salary


```excel
=MIN(C2:C6)
```

### Result

```text
25000
```

---

## What You Learned

1. How Excel identifies cells

2. How ranges work

```text
C2:C6
```

 SUM

 AVERAGE

 MAX

 MIN

---

## Real Business Intelligence Usage

Business Intelligence Analysts use these functions to:

- Calculate total revenue
- Calculate average order value
- Find highest sales month
- Find lowest performing region
- Build KPI dashboards
