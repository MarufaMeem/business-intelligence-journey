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

# Excel for Business Intelligence

This repository documents my Excel learning journey as part of my Business Intelligence Analyst roadmap.

The goal was to learn how businesses organize, analyze, summarize, and visualize data before moving to SQL and Power BI.

---

# Dataset Used

| Employee | Department | Salary |
|-----------|------------|----------|
| Rahim | Sales | 25000 |
| Karim | HR | 30000 |
| Jannat | Sales | 40000 |
| Nila | IT | 50000 |
| Rafi | HR | 35000 |

---

# Project 01 - Excel Basics

## Objective

Learn the fundamentals of Excel formulas and salary analysis.

## Screenshot

![Excel Basics](Excel/01_Excel_Basics_Screenshot.png)

## Functions Used

### SUM()

Calculates total salary.

```excel
=SUM(C2:C6)
```

Output:

```text
180000
```

Business Use:

- Total payroll calculation
- Revenue calculation
- Expense tracking

---

### AVERAGE()

Calculates average salary.

```excel
=AVERAGE(C2:C6)
```

Output:

```text
36000
```

Business Use:

- Average salary
- Average sales
- Average performance metrics

---

### MAX()

Returns highest salary.

```excel
=MAX(C2:C6)
```

Output:

```text
50000
```

Business Use:

- Highest sales
- Top performer identification

---

### MIN()

Returns lowest salary.

```excel
=MIN(C2:C6)
```

Output:

```text
25000
```

Business Use:

- Lowest salary analysis
- Minimum sales tracking

---

# Project 02 - IF Function

## Objective

Categorize employees based on salary.

## Screenshot

![IF Function](Excel/02_IF_Function_Screenshot.png)

## Formula

```excel
=IF(C2>30000,"High","Low")
```

## Output

| Employee | Salary | Salary Level |
|----------|----------|-------------|
| Rahim | 25000 | Low |
| Karim | 30000 | Low |
| Jannat | 40000 | High |
| Nila | 50000 | High |
| Rafi | 35000 | High |

## What I Learned

- Conditional logic
- Automated classification
- Business rule implementation

## Real Business Use

Companies use IF statements for:

- Employee grading
- Bonus eligibility
- Risk classification
- Customer segmentation

---

# Project 03 - COUNTIF Function

## Objective

Generate department-wise employee reports.

## Screenshot

![COUNTIF](Excel/03_COUNTIF_Screenshot.png)

## Formulas

### HR Employees

```excel
=COUNTIF(B2:B6,"HR")
```

Output:

```text
2
```

### Sales Employees

```excel
=COUNTIF(B2:B6,"Sales")
```

Output:

```text
2
```

### IT Employees

```excel
=COUNTIF(B2:B6,"IT")
```

Output:

```text
1
```

## Department Report

| Department | Employee Count |
|------------|----------------|
| HR | 2 |
| Sales | 2 |
| IT | 1 |

## What I Learned

- Conditional counting
- Workforce reporting
- Department analysis

## Business Use

COUNTIF is widely used for:

- Attendance reports
- Employee counts
- Sales transaction counts
- Inventory tracking

---

# Project 04 - PivotTable

## Objective

Summarize large datasets quickly without complex formulas.

## Screenshot

![PivotTable](Excel/04_PivotTable_Screenshot.png)

## Steps to Create

### Step 1

Select the dataset.

### Step 2

Click:

```text
Insert
→ PivotTable
```

### Step 3

Choose:

```text
New Worksheet
```

Click:

```text
OK
```

### Step 4

Drag fields:

Rows:

```text
Department
```

Values:

```text
Salary
```

Excel automatically creates a summarized report.

## Output

| Department | Total Salary |
|------------|-------------|
| HR | 65000 |
| IT | 50000 |
| Sales | 65000 |
| Grand Total | 180000 |

## What I Learned

- Data aggregation
- Dynamic reporting
- Business summarization

## Why PivotTables Matter

Without PivotTables:

- Manual calculations
- Repeated formulas
- Slower reporting

With PivotTables:

- Faster reporting
- Dynamic analysis
- Automatic summarization

---

# Project 05 - KPI Dashboard

## Objective

Create a simple Business Intelligence dashboard using Excel.

## Screenshot

![KPI Dashboard](Excel/05_KPI_Dashboard_Screenshot.png)

## KPI Metrics

### Total Employees

```excel
=COUNTA(A2:A6)
```

Output:

```text
5
```

### Total Salary

```excel
=SUM(C2:C6)
```

Output:

```text
180000
```

### Average Salary

```excel
=AVERAGE(C2:C6)
```

Output:

```text
36000
```

### Highest Salary

```excel
=MAX(C2:C6)
```

Output:

```text
50000
```

### Lowest Salary

```excel
=MIN(C2:C6)
```

Output:

```text
25000
```

## Dashboard Creation Steps

### Step 1

Create KPI table.

### Step 2

Select KPI values.

### Step 3

Click:

```text
Insert
→ Column Chart
```

### Step 4

Customize:

- Chart Title
- Data Labels
- Font Size
- Layout

## Dashboard Purpose

The dashboard provides an instant summary of:

- Employee Count
- Total Salary
- Average Salary
- Highest Salary
- Lowest Salary

---

# Excel Skills Learned

## Mathematical Functions

```excel
SUM()
AVERAGE()
MAX()
MIN()
COUNT()
COUNTIF()
```

## Logical Functions

```excel
IF()
```

## Reporting Tools

```text
PivotTables
Charts
KPI Dashboards
```

## Business Intelligence Concepts

- Data Analysis
- Reporting
- KPI Monitoring
- Dashboard Development
- Data Aggregation
- Decision Support

---

# Learning Outcome

After completing these projects, I can:

- Analyze employee datasets
- Create business reports
- Build KPI dashboards
- Summarize data using PivotTables
- Apply conditional logic
- Generate department-level insights
- Prepare data for SQL and Power BI

---

# SQL Business Intelligence Journey

This repository documents my SQL learning journey for Business Intelligence and Data Analytics using MySQL.

The goal is to learn how BI Analysts use SQL to:

- Store business data
- Retrieve information
- Analyze performance
- Build KPI reports
- Clean data
- Create reusable reports
- Prepare datasets for dashboards

Tools Used:

- MySQL Server
- MySQL Workbench

Database:

```sql
business_intelligence_journey
```

---

# Project 06 - Database and Tables

## Objective

Learn how to create databases, tables, and insert data.

## Create Database

```sql
CREATE DATABASE business_intelligence_journey;
USE business_intelligence_journey;
```

## Create Employees Table

```sql
CREATE TABLE employees(
employee_id INT PRIMARY KEY,
employee_name VARCHAR(50),
department VARCHAR(50),
salary INT
);
```

## Insert Data

```sql
INSERT INTO employees VALUES
(1,'rahim','sales',25000),
(2,'karim','hr',30000),
(3,'jannat','sales',40000),
(4,'nila','it',50000),
(5,'Rafi','HR',35000);
```

## Result

| employee_id | employee_name | department | salary |
|------------|--------------|------------|---------|
| 1 | rahim | sales | 25000 |
| 2 | karim | hr | 30000 |
| 3 | jannat | sales | 40000 |
| 4 | nila | it | 50000 |
| 5 | Rafi | HR | 35000 |

## Skills Learned

- CREATE DATABASE
- USE
- CREATE TABLE
- INSERT INTO

---

# Project 07 - Basic Queries

## Objective

Learn how to retrieve and sort data.

### Show All Employees

```sql
SELECT employee_name,salary
FROM employees;
```

### Result

| employee_name | salary |
|--------------|---------|
| rahim | 25000 |
| karim | 30000 |
| jannat | 40000 |
| nila | 50000 |
| Rafi | 35000 |

---

### Employees With Salary Greater Than 30000

```sql
SELECT *
FROM employees
WHERE salary > 30000;
```

### Result

| employee_name | salary |
|--------------|---------|
| jannat | 40000 |
| nila | 50000 |
| Rafi | 35000 |

---

### Highest Salary First

```sql
SELECT *
FROM employees
ORDER BY salary DESC;
```

### Result

| employee_name | salary |
|--------------|---------|
| nila | 50000 |
| jannat | 40000 |
| Rafi | 35000 |
| karim | 30000 |
| rahim | 25000 |

## Skills Learned

- SELECT
- WHERE
- ORDER BY

---

# Project 08 - Aggregations

## Objective

Learn how to calculate business metrics.

### Total Employees

```sql
SELECT COUNT(*) AS total_employees
FROM employees;
```

### Result

```text
5
```

---

### Total Salary

```sql
SELECT SUM(salary) AS total_salary
FROM employees;
```

### Result

```text
180000
```

---

### Average Salary

```sql
SELECT AVG(salary) AS average_salary
FROM employees;
```

### Result

```text
36000
```

---

### Employee Count by Department

```sql
SELECT department,
COUNT(*) AS employee_count
FROM employees
GROUP BY department;
```

### Result

| department | employee_count |
|------------|---------------|
| sales | 2 |
| hr | 1 |
| HR | 1 |
| it | 1 |

---

### Total Salary by Department

```sql
SELECT
department,
SUM(salary) AS total_salary
FROM employees
GROUP BY department;
```

### Result

| department | total_salary |
|------------|-------------|
| sales | 65000 |
| hr | 30000 |
| HR | 35000 |
| it | 50000 |

## Skills Learned

- COUNT
- SUM
- AVG
- GROUP BY

---

# Project 09 - Filtering Data

## Objective

Learn business filtering using WHERE, AND, and OR.

### IT Employees

```sql
SELECT *
FROM employees
WHERE department='it';
```

### Result

| employee_name | department |
|--------------|------------|
| nila | it |

---

### IT Employees With Salary Above 40000

```sql
SELECT *
FROM employees
WHERE department='IT'
AND salary > 40000;
```

### Result

| employee_name | department | salary |
|--------------|------------|---------|
| nila | it | 50000 |

---

### IT or HR Employees

```sql
SELECT *
FROM employees
WHERE department='it'
OR department='hr';
```

### Result

| employee_name | department |
|--------------|------------|
| karim | hr |
| nila | it |

## Skills Learned

- WHERE
- AND
- OR

---

# Project 10 - SQL JOINs

## Objective

Learn how to combine data from multiple tables.

## Create Department Table

```sql
CREATE TABLE departments(
department_id INT PRIMARY KEY,
department_name VARCHAR(50)
);
```

```sql
INSERT INTO departments VALUES
(1,'sales'),
(2,'hr'),
(3,'it');
```

---

### Employee Department Report

```sql
SELECT
e.employee_name,
d.department_name,
e.salary
FROM employees e
INNER JOIN departments d
ON e.department_id=d.department_id;
```

### Result

| employee_name | department_name | salary |
|--------------|----------------|---------|
| rahim | sales | 25000 |
| karim | hr | 30000 |
| jannat | sales | 40000 |
| nila | it | 50000 |

## Skills Learned

- INNER JOIN
- Primary Keys
- Foreign Keys

---

# Project 11 - HAVING

## Objective

Filter grouped results.

```sql
SELECT
d.department_name,
COUNT(*) AS employee_count,
SUM(e.salary) AS total_salary,
AVG(e.salary) AS average_salary
FROM employees e
INNER JOIN departments d
ON e.department_id=d.department_id
GROUP BY d.department_name
HAVING SUM(e.salary) > 60000;
```

### Result

| department_name | total_salary |
|----------------|-------------|
| sales | 65000 |

## Skills Learned

- HAVING
- Aggregate Filtering

---

# Project 12 - CASE Statements

## Objective

Create business classifications.

```sql
SELECT
employee_name,
salary,
CASE
WHEN salary > 30000 THEN 'High'
ELSE 'Low'
END AS salary_level
FROM employees;
```

### Result

| employee_name | salary | salary_level |
|--------------|---------|-------------|
| rahim | 25000 | Low |
| karim | 30000 | Low |
| jannat | 40000 | High |
| nila | 50000 | High |
| Rafi | 35000 | High |

## Skills Learned

- CASE
- Business Rules
- Classification

---

# Project 13 - Subqueries

## Objective

Compare values against calculated metrics.

```sql
SELECT
employee_name,
salary
FROM employees
WHERE salary >
(
SELECT AVG(salary)
FROM employees
);
```

### Result

| employee_name | salary |
|--------------|---------|
| jannat | 40000 |
| nila | 50000 |

## Skills Learned

- Subqueries
- Nested Queries

---

# Project 14 - Window Functions

## Objective

Rank employees by salary.

```sql
SELECT
employee_name,
salary,
ROW_NUMBER() OVER(
ORDER BY salary DESC
) AS salary_rank
FROM employees;
```

### Result

| employee_name | salary | salary_rank |
|--------------|---------|-------------|
| nila | 50000 | 1 |
| jannat | 40000 | 2 |
| Rafi | 35000 | 3 |
| karim | 30000 | 4 |
| rahim | 25000 | 5 |

## Skills Learned

- ROW_NUMBER()
- Window Functions
- Ranking

---

# Project 15 - Date Functions

## Objective

Analyze data by month and year.

```sql
SELECT
MONTH(sale_date) AS sale_month,
SUM(sale_amount) AS total_sales
FROM sales
GROUP BY MONTH(sale_date)
ORDER BY sale_month;
```

### Result

| Month | Total Sales |
|--------|------------|
| 1 | 24000 |
| 2 | 32000 |
| 3 | 35000 |
| 4 | 10000 |

## Skills Learned

- YEAR()
- MONTH()
- MONTHNAME()

---

# Project 16 - Date Filtering

## Objective

Filter data using date ranges.

```sql
SELECT
sale_id,
sale_amount,
sale_date
FROM sales
WHERE sale_date BETWEEN
'2026-01-01'
AND
'2026-02-28';
```

### Result

```text
6 sales transactions returned
```

## Skills Learned

- BETWEEN
- Date Filtering

---

# Project 17 - Data Cleaning

## Objective

Handle missing values.

```sql
SELECT
e.employee_name,
COALESCE(
d.department_name,
'Unknown Department'
) AS department_name
FROM employees e
LEFT JOIN departments d
ON e.department_id=d.department_id;
```

## Skills Learned

- NULL
- COALESCE
- LEFT JOIN
- Data Cleaning

---

# Project 18 - Data Quality Checks

## Objective

Identify unique values and data quality issues.

```sql
SELECT DISTINCT department
FROM employees;
```

```sql
SELECT
UPPER(department),
COUNT(*)
FROM employees
GROUP BY UPPER(department);
```

## Skills Learned

- DISTINCT
- UPPER()
- Data Quality Checks

---

# Project 19 - Common Table Expressions (CTE)

## Objective

Create reusable intermediate query results.

```sql
WITH average_salary AS (
SELECT AVG(salary) AS avg_salary
FROM employees
)
SELECT
employee_name,
salary
FROM employees
WHERE salary >
(
SELECT avg_salary
FROM average_salary
);
```

## Skills Learned

- WITH
- CTE
- Query Simplification

---

# Project 20 - SQL Views

## Objective

Create reusable reports.

```sql
CREATE VIEW employee_report AS
SELECT
e.employee_name,
d.department_name,
e.salary
FROM employees e
INNER JOIN departments d
ON e.department_id=d.department_id;
```

Use:

```sql
SELECT *
FROM employee_report;
```

## Skills Learned

- CREATE VIEW
- Reusable Reports

---

# Project 21 - Stored Procedures

## Objective

Automate SQL reporting.

```sql
DELIMITER //

CREATE PROCEDURE employee_report()
BEGIN
SELECT *
FROM employees;
END //

DELIMITER ;
```

Run:

```sql
CALL employee_report();
```

## Skills Learned

- CREATE PROCEDURE
- CALL
- Parameters
- SQL Automation

---

# Final Skills Acquired

By completing this SQL journey, I learned:

- Database Design
- Data Retrieval
- Data Filtering
- Aggregations
- Business Metrics
- Joins
- Data Cleaning
- Date Analysis
- Window Functions
- Subqueries
- CTEs
- Views
- Stored Procedures

These are core SQL skills required for Business Intelligence, Data Analyst, Reporting Analyst, and Junior BI Developer roles.

# Project 22 - UNION and UNION ALL

## UNION

Combines results and removes duplicate rows.

### Query

```sql
SELECT department
FROM employees

UNION

SELECT department
FROM contract_employees;
```

### Output

| Department |
|------------|
| sales |
| hr |
| HR |
| it |
| marketing |

### Main Info

- Combines results from multiple SELECT statements
- Removes duplicate rows
- Slower than UNION ALL
- MySQL must compare rows and remove duplicates before returning results

---

## UNION ALL

Combines results and keeps duplicate rows.

### Query

```sql
SELECT department
FROM employees

UNION ALL

SELECT department
FROM contract_employees;
```

### Output

| Department |
|------------|
| sales |
| hr |
| sales |
| it |
| HR |
| sales |
| marketing |
| it |

### Main Info

- Combines results from multiple SELECT statements
- Keeps duplicate rows
- Faster than UNION
- MySQL returns all rows directly without checking for duplicates

---

### Why is UNION ALL Faster?

UNION performs:

1. Combine rows
2. Compare rows
3. Remove duplicates
4. Return results

UNION ALL performs:

1. Combine rows
2. Return results

Since UNION ALL skips duplicate checking, it requires less processing and is usually faster on large datasets.

---
# Project 23 - String Functions

## UPPER()

Converts text to uppercase.

### Query

```sql
SELECT
employee_name,
UPPER(employee_name) AS uppercase_name
FROM employees;
```

### Output

| employee_name | uppercase_name |
|---------------|---------------|
| rahim | RAHIM |

### Main Info

- Converts text to uppercase
- Useful for standardizing text values
- Helps create consistent reports

---

## LOWER()

Converts text to lowercase.

### Query

```sql
SELECT
department,
LOWER(department) AS clean_department
FROM employees;
```

### Output

| department | clean_department |
|------------|------------------|
| HR | hr |

### Main Info

- Converts text to lowercase
- Helps clean inconsistent data
- Common in data preparation

---

## LENGTH()

Returns the number of characters in a string.

### Query

```sql
SELECT
employee_name,
LENGTH(employee_name) AS character_count
FROM employees;
```

### Output

| employee_name | character_count |
|---------------|----------------|
| rahim | 5 |

### Main Info

- Counts characters in text
- Useful for validation checks
- Helps identify unusually short or long values

---

## CONCAT()

Combines multiple text values.

### Query

```sql
SELECT
CONCAT(employee_name,' works in ',department)
AS employee_info
FROM employees;
```

### Output

| employee_info |
|--------------|
| rahim works in sales |

### Main Info

- Joins text together
- Useful for labels and reports
- Creates readable descriptions

---

## TRIM()

Removes extra spaces from text.

### Query

```sql
SELECT
TRIM('    Business Intelligence    ')
AS cleaned_text;
```

### Output

| cleaned_text |
|-------------|
| Business Intelligence |

### Main Info

- Removes leading and trailing spaces
- Helps clean imported data
- Improves data consistency

---

## LEFT()

Returns characters from the left side of a string.

### Query

```sql
SELECT
employee_name,
LEFT(employee_name,3) AS first_three_letters
FROM employees;
```

### Output

| employee_name | first_three_letters |
|---------------|--------------------|
| rahim | rah |

### Main Info

- Extracts characters from the beginning
- Useful for codes and abbreviations
- Common in text analysis

---

## RIGHT()

Returns characters from the right side of a string.

### Query

```sql
SELECT
employee_name,
RIGHT(employee_name,2) AS last_two_letters
FROM employees;
```

### Output

| employee_name | last_two_letters |
|---------------|----------------|
| rahim | im |

### Main Info

- Extracts characters from the end
- Useful for suffixes and IDs
- Common in data transformation

---

## REPLACE()

Replaces one text value with another.

### Query

```sql
SELECT
department,
REPLACE(department,'HR','Human Resources')
AS new_department
FROM employees;
```

### Output

| department | new_department |
|------------|---------------|
| HR | Human Resources |

### Main Info

- Replaces specific text
- Useful for data cleaning
- Helps standardize naming conventions

---
# Project 24 - Date Functions

## CURRENT_DATE()

Returns the current date.

### Query

```sql
SELECT CURRENT_DATE();
```

### Output

| CURRENT_DATE() |
|---------------|
| 2026-08-12 |

### Main Info

- Returns today's date
- No parameters required
- Useful for reports and dashboards

---

## CURRENT_TIMESTAMP()

Returns the current date and time.

### Query

```sql
SELECT CURRENT_TIMESTAMP();
```

### Output

| CURRENT_TIMESTAMP() |
|--------------------|
| 2026-08-12 09:15:30 |

### Main Info

- Returns current date and time
- Useful for audit logs
- Common in transactional systems

---

## DATEDIFF()

Calculates the difference between two dates.

### Query

```sql
SELECT
DATEDIFF('2026-12-31','2026-01-01')
AS total_days;
```

### Output

| total_days |
|-----------|
| 364 |

### Main Info

- Returns number of days between two dates
- End date comes first
- Useful for calculating durations

---

## DATE_ADD()

Adds time to a date.

### Query

```sql
SELECT
sale_date,
DATE_ADD(
sale_date,
INTERVAL 30 DAY
) AS next_review_date
FROM sales;
```

### Output

| sale_date | next_review_date |
|------------|------------------|
| 2026-01-10 | 2026-02-09 |

### Main Info

- Adds days, months, or years
- Useful for deadlines
- Common in scheduling systems

---

## DATE_SUB()

Subtracts time from a date.

### Query

```sql
SELECT
sale_date,
DATE_SUB(
sale_date,
INTERVAL 7 DAY
) AS previous_week
FROM sales;
```

### Output

| sale_date | previous_week |
|------------|--------------|
| 2026-01-10 | 2026-01-03 |

### Main Info

- Subtracts days, months, or years
- Useful for historical analysis
- Common in reporting filters

---

## Why Date Functions Matter

Business reports often answer questions like:

- How many days remain until a deadline?
- What were last month's sales?
- When should a follow-up occur?
- How long did a process take?

Date functions make these calculations easy and efficient.

---

# Project 25 - RANK() and DENSE_RANK()

## Objective

Learn how to rank data using SQL Window Functions.

Ranking is commonly used in:

- Sales Leaderboards
- Employee Performance Reports
- KPI Dashboards
- Top Customers Analysis

---

## RANK()

Assigns a rank to each row.

If duplicate values exist, they receive the same rank.

The next rank is skipped.

### Example

```text
1
2
3
3
5
6
7
```

### Query

```sql
SELECT
employee_name,
salary,
RANK() OVER(
ORDER BY salary DESC
) AS salary_rank
FROM employees;
```

### Example Output

| Employee | Salary | Rank |
|-----------|---------|------|
| Nila | 50000 | 1 |
| Jannat | 40000 | 2 |
| Rafi | 35000 | 3 |
| Karim | 35000 | 3 |
| Rahim | 25000 | 5 |

### Main Info

- Duplicate values share the same rank
- Next rank is skipped
- Useful for competition-style rankings

---

## DENSE_RANK()

Assigns a rank to each row.

If duplicate values exist, they receive the same rank.

No ranks are skipped.

### Example

```text
1
2
3
3
4
5
6
```

### Query

```sql
SELECT
employee_name,
salary,
DENSE_RANK() OVER(
ORDER BY salary DESC
) AS salary_rank
FROM employees;
```

### Example Output

| Employee | Salary | Rank |
|-----------|---------|------|
| Nila | 50000 | 1 |
| Jannat | 40000 | 2 |
| Rafi | 35000 | 3 |
| Karim | 35000 | 3 |
| Rahim | 25000 | 4 |

### Main Info

- Duplicate values share the same rank
- No gaps in ranking
- Useful for reporting and dashboards

---

## Difference

| Feature | RANK() | DENSE_RANK() |
|----------|----------|----------|
| Duplicate Values Share Rank | ✅ | ✅ |
| Skips Rank Numbers | ✅ | ❌ |
| Continuous Ranking | ❌ | ✅ |

---

## When to Use

### Use RANK()

When ranking positions matter.

Example:

```text
1st Place
2nd Place
3rd Place
3rd Place
5th Place
```

---

### Use DENSE_RANK()

When continuous numbering is preferred.

Example:

```text
1
2
3
3
4
```

---

## Business Intelligence Use Cases

- Top Performing Employees
- Highest Revenue Customers
- Product Rankings
- Sales Leaderboards
- Performance Dashboards

---

# Project 26 - LAG() and LEAD()

## Objective

Learn how to compare current values with previous and next values.

These functions are commonly used in:

- Sales Trend Analysis
- KPI Dashboards
- Growth Reports
- Month-over-Month Analysis
- Revenue Tracking

---

## LAG()

Returns the previous row value.

### Query

```sql
WITH monthly_sales AS (
SELECT
MONTH(sale_date) AS sale_month,
SUM(sale_amount) AS total_sales
FROM sales
GROUP BY MONTH(sale_date)
)

SELECT
sale_month,
total_sales,
LAG(total_sales)
OVER(
ORDER BY sale_month
) AS previous_month_sales
FROM monthly_sales;
```

### Output

| Month | Total Sales | Previous Month Sales |
|---------|------------|---------------------|
| 1 | 24000 | NULL |
| 2 | 32000 | 24000 |
| 3 | 35000 | 32000 |
| 4 | 10000 | 35000 |

### Main Info

- Looks backward one row
- Returns previous value
- First row returns NULL
- Useful for trend analysis

### Example

For Month 3:

```text
Current Month Sales = 35000

Previous Month Sales = 32000
```

LAG() retrieved the previous month's value.

---

## LEAD()

Returns the next row value.

### Query

```sql
WITH monthly_sales AS (
SELECT
MONTH(sale_date) AS sale_month,
SUM(sale_amount) AS total_sales
FROM sales
GROUP BY MONTH(sale_date)
)

SELECT
sale_month,
total_sales,
LEAD(total_sales)
OVER(
ORDER BY sale_month
) AS next_month_sales
FROM monthly_sales;
```

### Output

| Month | Total Sales | Next Month Sales |
|---------|------------|------------------|
| 1 | 24000 | 32000 |
| 2 | 32000 | 35000 |
| 3 | 35000 | 10000 |
| 4 | 10000 | NULL |

### Main Info

- Looks forward one row
- Returns next value
- Last row returns NULL
- Useful for forecasting and comparisons

---

## Monthly Growth Analysis

### Business Question

How much did sales increase compared to the previous month?

### Query

```sql
WITH monthly_sales AS (
SELECT
MONTH(sale_date) AS sale_month,
SUM(sale_amount) AS total_sales
FROM sales
GROUP BY MONTH(sale_date)
)

SELECT
sale_month,
total_sales,

LAG(total_sales)
OVER(
ORDER BY sale_month
) AS previous_month_sales,

total_sales -
LAG(total_sales)
OVER(
ORDER BY sale_month
) AS sales_growth

FROM monthly_sales;
```

### Output

| Month | Sales | Previous Month | Growth |
|---------|--------|----------------|---------|
| 1 | 24000 | NULL | NULL |
| 2 | 32000 | 24000 | 8000 |
| 3 | 35000 | 32000 | 3000 |
| 4 | 10000 | 35000 | -25000 |

### Main Info

- Positive value = Growth
- Negative value = Decline
- Used in KPI dashboards
- Common interview topic

---

## Difference

| Function | Purpose |
|-----------|---------|
| LAG() | Previous Row |
| LEAD() | Next Row |

# Project 27 - Running Totals

## Objective

Learn how to calculate cumulative totals using Window Functions.

Running Totals are commonly used in:

- Sales Dashboards
- Revenue Tracking
- KPI Reports
- Financial Analysis
- Executive Dashboards

---

## What is a Running Total?

A Running Total continuously adds current values to previous values.

Example:

| Month | Sales |
|---------|---------|
| Jan | 24000 |
| Feb | 32000 |
| Mar | 35000 |
| Apr | 10000 |

Running Total:

| Month | Running Total |
|---------|--------------|
| Jan | 24000 |
| Feb | 56000 |
| Mar | 91000 |
| Apr | 101000 |

---

## Running Total using SUM() OVER()

### Query

```sql
WITH monthly_sales AS (
SELECT
MONTH(sale_date) AS sale_month,
SUM(sale_amount) AS total_sales
FROM sales
GROUP BY MONTH(sale_date)
)

SELECT
sale_month,
total_sales,

SUM(total_sales)
OVER(
ORDER BY sale_month
) AS running_total

FROM monthly_sales;
```

### Output

| Month | Sales | Running Total |
|---------|---------|--------------|
| 1 | 24000 | 24000 |
| 2 | 32000 | 56000 |
| 3 | 35000 | 91000 |
| 4 | 10000 | 101000 |

---

## Main Info

- Calculates cumulative totals
- Uses Window Functions
- Does not collapse rows like GROUP BY
- Very common in dashboards and reports

---

## How It Works

Month 1:

```text
24000
```

Month 2:

```text
24000 + 32000
= 56000
```

Month 3:

```text
56000 + 35000
= 91000
```

Month 4:

```text
91000 + 10000
= 101000
```

---

## Running Average

### Query

```sql
WITH monthly_sales AS (
SELECT
MONTH(sale_date) AS sale_month,
SUM(sale_amount) AS total_sales
FROM sales
GROUP BY MONTH(sale_date)
)

SELECT
sale_month,
total_sales,

AVG(total_sales)
OVER(
ORDER BY sale_month
) AS running_average

FROM monthly_sales;
```

### Output

| Month | Sales | Running Average |
|---------|---------|----------------|
| 1 | 24000 | 24000 |
| 2 | 32000 | 28000 |
| 3 | 35000 | 30333 |
| 4 | 10000 | 25250 |

---

## Business Intelligence Use Cases

- Revenue Tracking
- Sales Monitoring
- Budget Analysis
- KPI Dashboards
- Financial Reporting

---

## Why Use Running Totals?

Without Running Total:

```text
Jan = 24000
Feb = 32000
Mar = 35000
Apr = 10000
```

With Running Total:

```text
Jan = 24000
Feb = 56000
Mar = 91000
Apr = 101000
```

This helps management see overall growth over time.

---
# Project 28 - KPI Calculations & Business Metrics

## Objective

Learn how Business Intelligence Analysts calculate KPIs used in dashboards and management reports.

These metrics are commonly used in:

- Power BI Dashboards
- Sales Reports
- Executive Reports
- HR Analytics
- Performance Tracking

---

# KPI 1 - Revenue Share %

## Business Question

What percentage of total sales came from each employee?

### Query

```sql
SELECT
e.employee_name,
SUM(s.sale_amount) AS employee_sales,

ROUND(
(
SUM(s.sale_amount)
/
(
SELECT SUM(sale_amount)
FROM sales
)
) * 100,
2
) AS revenue_share_percent

FROM sales s
INNER JOIN employees e
ON s.employee_id = e.employee_id

GROUP BY e.employee_name;
```

### Output

| Employee | Sales | Revenue Share % |
|-----------|---------|---------|
| Rahim | 23000 | 22.77 |
| Karim | 13000 | 12.87 |
| Jannat | 23000 | 22.77 |
| Nila | 33000 | 32.67 |
| Rafi | 9000 | 8.91 |

### Main Info

- Shows employee contribution to total sales
- Higher percentage means greater contribution
- Common KPI in sales dashboards

---

# KPI 2 - Monthly Growth %

## Business Question

How much did sales increase compared to the previous month?

### Query

```sql
WITH monthly_sales AS (

SELECT
MONTH(sale_date) AS sale_month,
SUM(sale_amount) AS total_sales

FROM sales

GROUP BY MONTH(sale_date)

)

SELECT

sale_month,
total_sales,

ROUND(
(
(total_sales -
LAG(total_sales)
OVER(ORDER BY sale_month))
/
LAG(total_sales)
OVER(ORDER BY sale_month)
) * 100,
2
) AS growth_percent

FROM monthly_sales;
```

### Output

| Month | Growth % |
|---------|---------|
| 1 | NULL |
| 2 | 33.33 |
| 3 | 9.38 |
| 4 | -71.43 |

### Main Info

- Positive value = Growth
- Negative value = Decline
- Measures month-over-month performance
- One of the most common business KPIs

---

# KPI 3 - Employee Performance Rating

## Business Question

Which employees are performing best?

### Query

```sql
SELECT

e.employee_name,

SUM(s.sale_amount) AS total_sales,

CASE

WHEN SUM(s.sale_amount) >= 30000
THEN 'Excellent'

WHEN SUM(s.sale_amount) >= 20000
THEN 'Good'

ELSE 'Needs Improvement'

END AS performance_rating

FROM sales s

INNER JOIN employees e
ON s.employee_id = e.employee_id

GROUP BY e.employee_name;
```

### Output

| Employee | Sales | Rating |
|-----------|---------|---------|
| Nila | 33000 | Excellent |
| Rahim | 23000 | Good |
| Jannat | 23000 | Good |
| Karim | 13000 | Needs Improvement |
| Rafi | 9000 | Needs Improvement |

### Main Info

- Uses CASE statement for classification
- Helps identify top performers
- Frequently used in HR and Sales dashboards

---

# KPI Summary Dashboard Query

### Query

```sql
SELECT

COUNT(*) AS total_employees,

SUM(salary) AS total_salary,

AVG(salary) AS average_salary,

MAX(salary) AS highest_salary,

MIN(salary) AS lowest_salary

FROM employees;
```

### Output

| KPI | Value |
|------|------|
| Total Employees | 5 |
| Total Salary | 180000 |
| Average Salary | 36000 |
| Highest Salary | 50000 |
| Lowest Salary | 25000 |

### Main Info

- Provides high-level business overview
- Commonly displayed using KPI cards in Power BI
- Frequently requested by management

---

# Why KPI Calculations Matter

KPIs help businesses answer questions like:

- Who generates the most revenue?
- Are sales growing or declining?
- Which employees perform best?
- Are business targets being achieved?

Without KPIs, management only sees raw data.

With KPIs, management can make decisions quickly.

---
# 📊 Power BI Complete Learning Guide

![Power BI](Power/power.png)


##  Table of Contents

1. [What is Power BI?](#-what-is-power-bi)
2. [Power Query (Data Transformation)](#-power-query-data-transformation)
3. [DAX Fundamentals](#-dax-fundamentals)
4. [DAX — Intermediate & Advanced](#-dax--intermediate--advanced)
5. [Data Modeling](#-data-modeling)
6. [Reports & Visuals](#-reports--visuals)
7. [Power BI Service (Cloud)](#-power-bi-service-cloud)
8. [How to Use This Repo](#-how-to-use-this-repo)
9. [Resources](#-resources)


---
##  What is Power BI?

Power BI is Microsoft's business intelligence tool used to:
- **Connect** to data sources (Excel, SQL, Web, APIs, etc.)
- **Transform** raw data into clean, usable tables (using **Power Query**)
- **Model** relationships between tables (Star Schema)
- **Analyze** data using formulas (using **DAX**)
- **Visualize** insights through interactive dashboards and reports

Think of it as: **Get Data → Clean Data (Power Query) → Model Data → Calculate (DAX) → Visualize → Publish (Service)**

---

##  Power Query (Data Transformation)

Power Query is the tool inside Power BI used to **clean and shape data before it hits your report**. It works like a recipe — every transformation step is recorded and can be replayed automatically every time data refreshes.

 Access it via: `Home → Transform Data`

### 1. Remove Columns
Deletes unnecessary columns to keep your data model clean and lightweight.
- Right-click a column → **Remove** (or **Remove Other Columns** to keep only selected ones)
-   *Tip: Fewer columns = faster reports.*

### 2. Rename Columns
Gives columns clear, business-friendly names.
- Double-click the column header, or right-click → **Rename**
-  *Tip: Avoid spaces/special characters if you'll reference the column in DAX often.*

### 3. Change Data Type
Ensures each column is stored correctly (Text, Whole Number, Decimal, Date, Boolean, etc.)
- Click the data type icon next to the column header
-  *Wrong data types are the #1 cause of DAX errors for beginners.*

### 4. Replace Values
Finds and replaces specific values in a column (e.g., replacing `"N/A"` with `null`, or fixing typos).
- Right-click column → **Replace Values**

### 5. Remove Duplicates
Removes duplicate rows based on selected column(s) to avoid double-counting.
- Right-click column(s) → **Remove Duplicates**
-  *Always check whether duplicates are true errors or valid repeated transactions before removing.*

### 6. Filter Rows
Keeps only the rows that meet certain conditions (e.g., only "2024" sales, exclude blanks).
- Click the dropdown arrow on any column header → apply filter conditions

### 7. Custom Column
Creates a brand-new column using a custom Power Query (M) formula.
- `Add Column → Custom Column`
```m
= [Price] * [Quantity]
```

### 8. Conditional Column
Creates a column based on if/then logic — no formula writing needed, built through a UI wizard.
- `Add Column → Conditional Column`
- Example: If `Sales > 1000` then `"High"` else `"Low"`

### 9. Split/Merge Columns
- **Split Column**: Breaks one column into multiple (e.g., splitting "First Last" into First Name / Last Name) using a delimiter.
- **Merge Columns**: Combines multiple columns into one (e.g., combining City + State).

### 10. Group By
Aggregates data — similar to a Pivot Table's grouping. Example: Total Sales grouped by Region.
- `Transform → Group By` → choose column to group by + aggregation (Sum, Count, Average, etc.)

### 11. Merge Queries
Power Query's version of a SQL **JOIN** — combines two tables based on a matching column.
- `Home → Merge Queries`
- Choose Join Kind: Inner, Left Outer, Right Outer, Full Outer, Anti Join

### 12. Append Queries
Stacks two or more tables **on top of each other** (like SQL `UNION`) — used when tables have the same columns but different rows (e.g., Jan sales + Feb sales).
- `Home → Append Queries`

### 13. Pivot/Unpivot
- **Unpivot**: Converts columns into rows — turns "wide" data into "long/tidy" data. Extremely common when cleaning Excel exports (e.g., unpivoting Jan, Feb, Mar columns into a single "Month" column).
- **Pivot**: The reverse — converts row values into columns.
- `Transform → Pivot Column` / `Unpivot Columns`

---

##  DAX Fundamentals

DAX is the formula language used in Power BI to build calculations. If Power Query is about **shaping** data, DAX is about **calculating and analyzing** it.

### Measures vs Calculated Columns
| | **Measure** | **Calculated Column** |
|---|---|---|
| Calculated | At query/visual time (dynamically) | At data refresh time (row-by-row) |
| Storage | Not stored, computed on the fly | Stored in the table, takes up memory |
| Context | Responds to filters/slicers | Fixed per row |
| Use case | Aggregations (Total Sales, Avg Price) | Row-level classification (e.g., Age Group) |
| Where used | Visuals, cards, KPIs | Can be used as a filter/axis/slicer field |

>  **Rule of thumb**: If you're aggregating (sum, average, count) → use a **Measure**. If you need a static per-row value → use a **Calculated Column**.

### `SUM()`
Adds up all values in a numeric column.
```dax
Total Sales = SUM(Sales[SalesAmount])
```

### `AVERAGE()`
Calculates the arithmetic mean of a column.
```dax
Avg Sales = AVERAGE(Sales[SalesAmount])
```

### `DISTINCTCOUNT()`
Counts the number of **unique** values in a column.
```dax
Unique Customers = DISTINCTCOUNT(Sales[CustomerID])
```

### `CALCULATE()`
The most important function in DAX. It changes the **filter context** of a calculation.
```dax
Sales 2024 = CALCULATE(SUM(Sales[SalesAmount]), Sales[Year] = 2024)
```
>  Almost every advanced DAX pattern (YoY growth, running totals, RLS) is built using `CALCULATE()`.

### `FILTER()`
Returns a filtered **table** (not a value) — usually used inside other functions like `CALCULATE()` or `SUMX()`.
```dax
High Value Sales =
CALCULATE(
    SUM(Sales[SalesAmount]),
    FILTER(Sales, Sales[SalesAmount] > 1000)
)
```

### `IF()`
Standard conditional logic — returns one value if a condition is true, another if false.
```dax
Sales Status = IF(SUM(Sales[SalesAmount]) > 5000, "High", "Low")
```

### `SWITCH()`
A cleaner alternative to writing multiple nested `IF()` statements.
```dax
Sales Category =
SWITCH(
    TRUE(),
    SUM(Sales[SalesAmount]) > 10000, "Excellent",
    SUM(Sales[SalesAmount]) > 5000, "Good",
    "Needs Improvement"
)
```

### `DIVIDE()`
A safe division function that automatically handles divide-by-zero errors (returns blank instead of an error).
```dax
Profit Margin = DIVIDE(SUM(Sales[Profit]), SUM(Sales[SalesAmount]), 0)
```

### `ALL()`
Removes filters from a table or column — useful for calculating "% of total" style metrics.
```dax
% of Total Sales =
DIVIDE(
    SUM(Sales[SalesAmount]),
    CALCULATE(SUM(Sales[SalesAmount]), ALL(Sales))
)
```

### `REMOVEFILTERS()`
A more explicit, modern alternative to `ALL()` — clears filters from specific columns/tables while making the intent clearer in your code.
```dax
Total Sales All Regions =
CALCULATE(SUM(Sales[SalesAmount]), REMOVEFILTERS(Sales[Region]))
```

### `RELATED()`
Pulls a value from a **related table** (used on the "many" side of a relationship) — similar to a VLOOKUP but relationship-driven.
```dax
Category Name = RELATED(Products[Category])
```

### `RELATEDTABLE()`
Returns a **table** of related rows from the "many" side, used from the "one" side of a relationship.
```dax
Number of Orders = COUNTROWS(RELATEDTABLE(Sales))
```

---

##  DAX — Intermediate & Advanced

### Row Context vs Filter Context
This is the single most important concept for understanding *why* DAX behaves the way it does.

- **Row Context**: Exists when DAX evaluates a formula **one row at a time** — like in a Calculated Column, or inside iterator functions (`SUMX`, `FILTER`). DAX "walks" through the table row by row.
- **Filter Context**: Exists when a set of filters (from slicers, visuals, or `CALCULATE`) narrows down which rows are visible **before** an aggregation happens. A Measure like `SUM(Sales[Amount])` on a Region visual is filtered by whatever Region is currently selected.

```dax
// Row context example — evaluated per row, then implicitly summed
Total Revenue Column = [Price] * [Quantity]

// Filter context example — depends on what's selected in the report
Total Revenue Measure = SUM(Sales[Price]) * SUM(Sales[Quantity])
```
>  *These two examples do NOT always give the same result — that's the whole point of understanding context.*

`CALCULATE()` is special because it can **transition row context into filter context** — this is called **context transition**, and it's why `CALCULATE()` is used so heavily inside calculated columns and iterators.

### `SUMX()` / Iterator Functions
Iterator functions (`SUMX`, `COUNTX`, `AVERAGEX`, `MAXX`, `MINX`, `RANKX`) evaluate an expression **row by row**, then aggregate the results. They're needed whenever a calculation can't be done with a simple column reference — e.g., multiplying two columns together per row before summing.
```dax
Total Revenue =
SUMX(
    Sales,
    Sales[Price] * Sales[Quantity]
)
```
> 💡 Use `SUMX` instead of a calculated column when you don't want the value stored — it calculates fresh every time based on current filters.

### `COUNTX()` / `AVERAGEX()`
Same iterator logic as `SUMX`, just with a different final aggregation.
```dax
Count of Profitable Orders =
COUNTX(
    FILTER(Sales, Sales[Profit] > 0),
    Sales[OrderID]
)

Average Order Value =
AVERAGEX(
    Sales,
    Sales[Price] * Sales[Quantity]
)
```

### `VALUES()`
Returns a **table** of distinct values from a column, respecting the current filter context. Commonly used inside `CALCULATE`, `FILTER`, or to check how many items are currently selected.
```dax
Distinct Products Selling = COUNTROWS(VALUES(Products[ProductName]))
```

### `SELECTEDVALUE()`
Returns the single selected value from a column when exactly one value is filtered (e.g., one item picked in a slicer) — otherwise returns a blank or a default you specify. Great for dynamic titles.
```dax
Selected Region = SELECTEDVALUE(Sales[Region], "All Regions")
```

### More Advanced `CALCULATE()` Patterns
`CALCULATE()` can take multiple filter arguments, and combine with modifier functions for precise control:

```dax
// Multiple conditions
Sales - East 2024 =
CALCULATE(
    SUM(Sales[SalesAmount]),
    Sales[Region] = "East",
    Sales[Year] = 2024
)

// KEEPFILTERS - adds to existing filters instead of overriding them
Sales High Value =
CALCULATE(
    SUM(Sales[SalesAmount]),
    KEEPFILTERS(Sales[SalesAmount] > 1000)
)

// ALLEXCEPT - removes all filters except the ones listed
% of Category Total =
DIVIDE(
    SUM(Sales[SalesAmount]),
    CALCULATE(SUM(Sales[SalesAmount]), ALLEXCEPT(Products, Products[Category]))
)
```

### Ranking with `RANKX()`
Ranks values in a column/measure against each other — great for "Top N Products" style visuals.
```dax
Product Rank =
RANKX(
    ALL(Products[ProductName]),
    [Total Sales],
    ,
    DESC
)
```

### Running Totals
Accumulates a value over time (e.g., cumulative sales through the year).
```dax
Running Total Sales =
CALCULATE(
    SUM(Sales[SalesAmount]),
    FILTER(
        ALL('Date'),
        'Date'[Date] <= MAX('Date'[Date])
    )
)
```

### Month-over-Month (MoM) Growth
```dax
Sales Previous Month =
CALCULATE(
    SUM(Sales[SalesAmount]),
    PREVIOUSMONTH('Date'[Date])
)

MoM Growth % =
DIVIDE(
    [Total Sales] - [Sales Previous Month],
    [Sales Previous Month]
)
```

### Year-over-Year (YoY) Growth
```dax
Sales Previous Year =
CALCULATE(
    SUM(Sales[SalesAmount]),
    SAMEPERIODLASTYEAR('Date'[Date])
)

YoY Growth % =
DIVIDE(
    [Total Sales] - [Sales Previous Year],
    [Sales Previous Year]
)
```

### Time Intelligence
A family of built-in DAX functions that make date-based calculations easy — but they **require a proper Date Table** marked as such (see below).

| Function | Purpose |
|---|---|
| `TOTALYTD()` | Year-to-date total |
| `TOTALQTD()` | Quarter-to-date total |
| `TOTALMTD()` | Month-to-date total |
| `DATESYTD()` | Returns a table of dates from year start to current date |
| `SAMEPERIODLASTYEAR()` | Shifts the current filter back exactly one year |
| `PARALLELPERIOD()` | Shifts a date period by a given number of months/quarters/years |
| `DATEADD()` | Flexible date shifting in either direction |

```dax
Sales YTD = TOTALYTD(SUM(Sales[SalesAmount]), 'Date'[Date])
```

---

##  Data Modeling

### Date Table
A dedicated table containing one row per calendar date, used to power all Time Intelligence functions. Without it, YoY/MoM/YTD formulas won't work correctly.
```dax
DateTable =
CALENDAR(DATE(2020,1,1), DATE(2026,12,31))
```
After creating it, go to `Table Tools → Mark as Date Table` and select the date column. Add helper columns like Year, Month Name, Quarter for easy slicing.

### Star Schema
The recommended way to structure a Power BI data model: **one central Fact table** surrounded by **Dimension tables**, connected by relationships — visually forming a "star" shape. It's faster, easier to maintain, and avoids duplicate/inconsistent data compared to one giant flat table.

### Fact vs Dimension Tables
- **Fact Table**: Contains transactional/numeric data that gets measured — e.g., `Sales` (SalesAmount, Quantity, OrderDate). Usually long and narrow (millions of rows, few columns).
- **Dimension Table**: Contains descriptive attributes used to filter/group facts — e.g., `Products`, `Customers`, `Date`. Usually short and wide (fewer rows, more columns).

> 💡 *Rule of thumb: if a column describes "who, what, where, when" → dimension. If it's a number you sum/count/average → fact.*

### Relationships
Links between tables based on a common column (usually a Dimension's primary key to a Fact's foreign key).
- Go to the **Model view** → drag a column from one table to the matching column in another to create a relationship
- Relationships have a **Cardinality**: One-to-Many (most common), One-to-One, or Many-to-Many

### Cross-Filter Direction
Controls which direction a filter can "flow" through a relationship.
- **Single direction**: Filters flow from the "one" side (dimension) to the "many" side (fact) — the default and safest choice.
- **Both (bidirectional)**: Filters flow both ways — powerful but can cause ambiguous filtering and performance issues if overused. Use only when needed (e.g., many-to-many bridge tables).

### Many-to-Many Relationships
Occurs when neither table has a unique key matching the other directly — e.g., a Customer can have many Bank Accounts, and an Account can be shared by many Customers. Solved by:
1. Creating a **bridge table** containing the unique combinations, or
2. Setting the relationship cardinality to "Many-to-Many" directly (supported in modern Power BI, but use carefully)

### Advanced DAX KPI Patterns
Combine measures to build reusable, dynamic KPI cards:
```dax
KPI Status =
VAR CurrentSales = [Total Sales]
VAR Target = [Sales Target]
RETURN
SWITCH(
    TRUE(),
    CurrentSales >= Target, "🟢 On Track",
    CurrentSales >= Target * 0.8, "🟡 At Risk",
    "🔴 Behind"
)
```
> 💡 `VAR`/`RETURN` blocks make complex DAX measures far easier to read and debug — always prefer them over deeply nested formulas.

---

## 🎨 Reports & Visuals

### Drill-Through
Lets users right-click a data point (e.g., a specific product) and jump to a detailed page filtered to just that context.
- Create a new report page → drag the relevant field into the **Drill through** filter well → right-click any visual value on another page → **Drill through**

### Tooltips
Custom mini report pages that appear when hovering over a visual, showing extra detail without cluttering the main view.
- Create a page → set **Page Information → Allow use as tooltip** → set its size to "Tooltip" → assign it in the **Format → Tooltip** pane of your visual

### Bookmarks
Captures the current state of a report page (filters, slicer selections, visibility of objects) so it can be recalled instantly — used heavily for building interactive "guided" dashboards and toggle views.
- `View → Bookmarks Pane → Add`

### Buttons & Page Navigation
Buttons combined with bookmarks or built-in navigation actions let users click through a report like an app instead of using page tabs.
- `Insert → Buttons` → set the **Action** (Page navigation, Bookmark, Back, Web URL, etc.)

### Conditional Formatting
Dynamically changes the color/icon of cells, bars, or backgrounds based on a value or measure — great for highlighting KPIs at a glance.
- Select a visual → `Format → Cell Elements` → choose a field → **Conditional Formatting** → Rules, Color Scale, or Data Bars

### Advanced Matrix/Table Visuals
Matrix visuals support drill-down hierarchies, subtotal control, and per-column formatting — much more powerful than a basic Table for multi-level reporting (e.g., Region → Country → City).
- Enable **Row/Column subtotals**, use the **+/-** icons for drill-down, and apply conditional formatting per measure column

### Dashboard Design Principles
- Keep the most important KPI top-left (where the eye naturally goes first)
- Use a consistent color palette — 2–3 core colors max, plus a highlight color
- Avoid clutter — every visual should answer a specific business question
- Use whitespace deliberately; don't fill every pixel
- Group related visuals together, and follow a logical left-to-right, top-to-bottom reading flow

---

##  Power BI Service (Cloud)

### Publish to Power BI Service
Once your report is built in Power BI Desktop, publish it to the cloud for sharing.
- `Home → Publish` → select a destination **Workspace**

### Workspaces
Cloud containers that organize related reports, dashboards, and datasets — similar to a folder, but with collaboration and access control built in. Types include personal workspaces and shared workspaces (for teams).

### Refresh
Keeps your published dataset up to date with the source data.
- **Manual refresh**: Click "Refresh" in the Service
- **Scheduled refresh**: Set up automatic refresh times in Dataset Settings (requires a **Gateway** if your data source is on-premises, like a local SQL Server or Excel file)

### Sharing
Multiple ways to distribute a report depending on audience:
- **Share**: Direct link to specific people
- **Apps**: Package multiple reports/dashboards into a polished, versioned "app" for a wider audience
- **Embed**: Embed a report into a website, Teams, or SharePoint
- **Publish to Web**: Public, unsecured embed (use with caution — no login required)

### Row-Level Security (RLS)
Restricts what data different users can see within the **same report** — e.g., a Regional Manager only sees their own region's sales.
1. In Power BI Desktop: `Modeling → Manage Roles` → create a role → add a DAX filter
```dax
[Region] = "East"
```
2. For **dynamic RLS** (auto-detects the logged-in user), use:
```dax
[UserEmail] = USERPRINCIPALNAME()
```
3. Publish, then in the Service go to **Dataset Settings → Security** to assign users to roles.

---


## 📖 Resources

- [Official Power BI Documentation](https://learn.microsoft.com/en-us/power-bi/)
- [DAX Guide (SQLBI)](https://dax.guide/)
- [Power Query M Function Reference](https://learn.microsoft.com/en-us/powerquery-m/)

---


