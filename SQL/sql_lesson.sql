create database business_intelligence_journey;
use  business_intelligence_journey;
create table employees( 
employee_id int primary key,
employee_name varchar(50),
department varchar(50),
salary int
);

insert into employees values(1,'rahim','sales',25000),
(2,'karim','hr',30000),
(3,'jannat','sales',40000),
(4,'nila','it',50000),
(5,'Rafi','HR',35000);

select * from employees where salary>30000;

select * from employees order by salary desc;

SELECT employee_name,salary
FROM employees;

SELECT COUNT(*) AS total_employees
FROM employees;

select sum(salary) as total_salary from employees;

SELECT AVG(salary) AS average_salary
FROM employees;

select department,count(*) as employee_count 
from employees group by department;

SELECT
department,
SUM(salary) AS total_salary
FROM employees
GROUP BY department;

SELECT
department,
AVG(salary) AS average_salary
FROM employees
GROUP BY department;


select * from employees where department='it';

SELECT *
FROM employees
WHERE department = 'IT'
AND salary > 40000;

select * from employees where department='it' 
or department='hr';

SELECT *
FROM employees
WHERE department = 'HR'
AND salary > 30000
ORDER BY salary DESC;





CREATE TABLE departments (
    department_id INT PRIMARY KEY,
    department_name VARCHAR(50)
);


insert into departments values(1,'sales'),(2,'hr'),(3,'it');

SELECT * FROM departments;


alter table employees add department_id int;


UPDATE employees
SET department_id = 1
WHERE department = 'sales';

UPDATE employees
SET department_id = 2
WHERE department = 'hr';

UPDATE employees
SET department_id = 3
WHERE department = 'it';

select e.employee_name,d.department_name,e.salary
from employees e
inner join departments d 
on e.department_id=d.department_id;


SET SQL_SAFE_UPDATES = 0;


SELECT
    e.employee_name,
    d.department_name,
    e.salary
FROM employees e
INNER JOIN departments d
    ON e.department_id = d.department_id
WHERE d.department_name = 'HR'
AND e.salary > 30000;


SELECT
    d.department_name,
    COUNT(*) AS employee_count
FROM employees e
INNER JOIN departments d
    ON e.department_id = d.department_id
GROUP BY d.department_name;

SELECT
    d.department_name,
    COUNT(*) AS employee_count,
    SUM(e.salary) AS total_salary,
    AVG(e.salary) AS average_salary
FROM employees e
INNER JOIN departments d
    ON e.department_id = d.department_id
GROUP BY d.department_name
HAVING SUM(e.salary) > 60000;

SELECT
    d.department_name
FROM employees e
INNER JOIN departments d
    ON e.department_id = d.department_id
GROUP BY d.department_name
HAVING avg(e.salary) > 30000;

SELECT
    employee_name,
    salary,
    CASE
        WHEN salary > 30000 THEN 'High'
        ELSE 'Low'
    END AS salary_level
FROM employees;


SELECT
    e.employee_name,
    d.department_name,
    e.salary,
    CASE
        WHEN e.salary >= 40000 THEN 'High'
        WHEN e.salary >= 30000 THEN 'Medium'
        ELSE 'Low'
    END AS salary_level
FROM employees e
INNER JOIN departments d
    ON e.department_id = d.department_id;
    
SELECT
    employee_name,
    salary
FROM employees
WHERE salary > (
    SELECT AVG(salary)
    FROM employees
);

SELECT
    employee_name,
    salary,
    ROW_NUMBER() OVER (
        ORDER BY salary DESC
    ) AS salary_rank
FROM employees;

SELECT
    e.employee_name,
    d.department_name,
    e.salary,
    ROW_NUMBER() OVER (
        PARTITION BY d.department_name
        ORDER BY e.salary DESC
    ) AS department_rank
FROM employees e
LEFT JOIN departments d
    ON e.department_id = d.department_id;
    
    
    
    
CREATE TABLE sales (
    sale_id INT PRIMARY KEY,
    employee_id INT,
    sale_amount INT,
    sale_date DATE
);
INSERT INTO sales
VALUES
(1, 1, 5000, '2026-01-10'),
(2, 2, 7000, '2026-01-15'),
(3, 3, 12000, '2026-01-20'),
(4, 1, 8000, '2026-02-05'),
(5, 4, 15000, '2026-02-12'),
(6, 5, 9000, '2026-02-20'),
(7, 3, 11000, '2026-03-03'),
(8, 2, 6000, '2026-03-15'),
(9, 4, 18000, '2026-03-25'),
(10, 1, 10000, '2026-04-10');
SELECT * FROM sales;

select sale_id,sale_date ,
year(sale_date) as sale_year
from sales;


SELECT
    MONTH(sale_date) AS sale_month,
    SUM(sale_amount) AS total_sales
FROM sales
GROUP BY MONTH(sale_date)
ORDER BY sale_month;

SELECT
    MONTHNAME(sale_date) AS month_name,
    COUNT(*) AS transaction_count,
    SUM(sale_amount) AS total_sales,
    AVG(sale_amount) AS average_sale
FROM sales
GROUP BY MONTH(sale_date), MONTHNAME(sale_date)
ORDER BY MONTH(sale_date);

SELECT
    sale_id,
    sale_amount,
    sale_date
FROM sales
WHERE sale_date BETWEEN '2026-01-01' AND '2026-02-28';

SELECT
    e.employee_name,
    COALESCE(d.department_name, 'Unknown Department') AS department_name,
    e.salary
FROM employees e
LEFT JOIN departments d
    ON e.department_id = d.department_id;
    
WITH average_salary AS (
    SELECT AVG(salary) AS avg_salary
    FROM employees
)
SELECT
    employee_name,
    salary
FROM employees
WHERE salary > (
    SELECT avg_salary
    FROM average_salary
);


WITH department_summary AS (
    SELECT
        d.department_name,
        COUNT(*) AS employee_count,
        SUM(e.salary) AS total_salary,
        AVG(e.salary) AS average_salary
    FROM employees e
    INNER JOIN departments d
        ON e.department_id = d.department_id
    GROUP BY d.department_name
)
SELECT
    department_name,
    employee_count,
    total_salary,
    average_salary
FROM department_summary
WHERE total_salary > 60000;


CREATE VIEW employee_report AS
SELECT
    e.employee_name,
    d.department_name,
    e.salary
FROM employees e
INNER JOIN departments d
    ON e.department_id = d.department_id;
    
    
    DELIMITER //

CREATE PROCEDURE employee_report()
BEGIN
    SELECT *
    FROM employees;
END //

DELIMITER ;
CALL employee_report();



DELIMITER //

CREATE PROCEDURE employees_by_department(
    IN dept_name VARCHAR(50)
)
BEGIN

    SELECT *
    FROM employees
    WHERE department = dept_name;

END //

DELIMITER ;
CALL employees_by_department('sales');

SHOW PROCEDURE STATUS
WHERE Db='business_intelligence_journey';

DROP PROCEDURE employee_report;