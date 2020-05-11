-- Fix birthdate vlues in Employees table, need to subtract 100 years from each value
-- Found how to perform this operation from the site:
-- https://www.postgresql.org/message-id/4199e4a414eb3dd2e9a25896f0693be6@biglumber.com
UPDATE Employees
SET birth_date = birth_date - '100 years'::interval;

-- List the following details of each employee: employee number, last name, first name, sex, and salary.
SELECT E.emp_no, E.last_name, E.first_name, E.sex, S.salary
FROM Employees AS E
JOIN Salaries AS S
ON (E.emp_no = S.emp_no);

-- List first name, last name, and hire date for employees who were hired in 1986.
-- Can use EXTRACT() to get year from a timestamp; found this information from the site:
-- https://www.postgresqltutorial.com/postgresql-extract/
SELECT first_name, last_name, hire_date
FROM Employees
WHERE EXTRACT(YEAR FROM hire_date) = 1986;