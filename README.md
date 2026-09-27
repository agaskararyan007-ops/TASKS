# TASKS
CONTAINS ALL THE TASKS PERFORMED AT IT VEDANT REGARDING ADV.EXCEL, SQL, POWER BI, etc.


# EXCEL TASKS

This repository documents the tasks and projects I've completed during my time at IT Vedant Institute, organized within the tasks folder for easy navigation.
It reflects my hands-on learning journey in AI and Data Science, covering practical exercises in Excel, data analysis, and problem-solving.
Each subfolder represents a different topic, making it simple to track progress over time.
More sections will be added as I continue learning and expanding my skill set.


# EXCEL_TASK-1

## Question1 – School Attendance Monitoring

SUM, MAX, MIN, COUNTA to total present/school days, find highest/lowest attendance and late days, and count students.
LARGE and SMALL to find the 2nd/3rd highest and lowest attendance values.
Named Ranges (Present_Days, Total_Days, Late_Days) used in place of raw cell references.
All of the above combined into a final attendance summary report.

## Question2 – Cell Referencing (Relative, Absolute, Mixed)

Relative referencing — Total Amount column recalculates row by row (Quantity × Unit Price).
Absolute referencing — a single locked discount-rate cell ($) applied across all orders.
Mixed referencing — menu prices calculated across Dine In, Take Away, and Delivery multipliers.

# EXCEL_TASK-2

## 📊Conditional Formatting — Employee Sales Highlighting

Dataset: 56 employee records (Name, Sales) with a built-in "Questions" column listing 9 conditional-formatting tasks to complete
---------------------------------------------------------------------------------------------------------------------------------
Solved using rule-based conditional formatting across the sheet:
-----------------------------------------------------------------------------------------------------------------------------
Formula rule (EXACT+UPPER) → highlights names written fully in uppercase
--
Text-contains rule → highlights names containing "SHARMA"
-
Cell value rule (≥ 6) → highlights employees with Sales of 6 or higher
-
Data Bars → in-cell bar visualization on the Sales column
-
Duplicate Values rule → flags repeated Sales figures
-
Top 10 rule (top 5) → highlights the 5 highest Sales values
-
Top 10 rule (bottom 5) → highlights the 5 lowest Sales values
-

# EXCEL_TASK-3

## 📊Contact Lookup & Data Aggregation

A solved Excel practice file demonstrating real-world lookup and aggregation formulas commonly used in data/CRM workflows.

🔧 What's inside

1. Contact Details Sheet
Retrieves a customer's Contact Number from a master Data sheet using their Lead ID, powered by XLOOKUP:

excel
=XLOOKUP(B5, Data!$E$3:$E$99, Data!$B$3:$B$99)

2. Data Sheet
The master reference table containing Lead ID, Contact Number, VAT ID, and Email — the source data for the lookups above.

3. Example Sheet — Individual Product Amount Calculator
A SUMIFS-based cross-tab that totals purchase amounts per person, per product category, using two-condition matching:

excel
=SUMIFS($D$5:$D$19, $B$5:$B$19, G5, $C$5:$C$19, H4)

Includes row/column totals via SUM.

🧠 Skills demonstrated
XLOOKUP for single-key lookups across sheets
SUMIFS for multi-criteria conditional aggregation
Building a simple pivot-style summary table from transactional (long-format) data
Structuring reference/master data separately from working sheets

# Note: the sheet lists 9 tasks total, but two are worded identically (items 2 and 9, both "Sales ≥ 6"), and I didn't find a distinct rule for item 5 ("names starting with A") 


# EXCEL_TASK-4

## 📊Dynamic Product Sales Dashboard (INDEX + MATCH)

A solved Excel practice file demonstrating a dynamic two-dimensional lookup using nested INDEX + MATCH — a classic alternative to VLOOKUP when data needs to be pulled by both row and column criteria.

🔧 What's inside

Sales Data Table (A1:Q8)
Quarterly sales figures for 6 products (Biscuits, Chocolates, Chocochips, Cookies, Nutri Bars, Jelly Beans) across 4 regions (East, West, North, South), each broken into QTR 1–4.

Interactive Summary Table (G10:K16)
A dropdown at I10 lets you select any product, and the table below automatically updates to show that product's quarterly sales across all 4 regions — powered by a double MATCH inside INDEX:

excel
=INDEX($B$3:$Q$8, MATCH($I$10,$A$3:$A$8,0), MATCH(H$12,$B$1:$Q$1,0)+MATCH($G13,$B$2:$Q$2,0)-1)
----------------------------------------------------------------------------------------------------------------------------------------------------------------
🧠 Skills demonstrated
INDEX + MATCH (nested/double match) for 2D lookups
Handling repeating column headers (quarters repeat per region block) by combining a region-block offset with a quarter offset
Building a dropdown-driven, dynamic summary table that updates live based on user selection
Array-style formula logic without relying on XLOOKUP/dynamic array spill functions


# EXCEL_TASK-5

## 🆔Employee Van ID Generator (Text Functions)

A solved Excel practice file demonstrating text manipulation and string concatenation to auto-generate unique employee IDs from HR data.

🔧 What's inside

Employee Table (C16:F36)
20 employee records (Employee ID, First Name, Last Name) with a computed Van ID column built using nested text functions:

excel
=UPPER("VI"&LEFT(D17,1)&RIGHT(E17,3)&C17)
📐 Business Rule

Each Van ID is built as:

"VI" — fixed prefix
first letter of First Name
last 3 letters of Last Name
Employee ID (appended as-is)
→ entire result converted to UPPERCASE

Example: Mateen Shaikh, ID 3342 → VIMIKH3342
----------------------------------------------------------------------------------------------------------------------------------------------------------------
🧠 Skills demonstrated
LEFT() and RIGHT() for substring extraction
String concatenation with &
UPPER() for case normalization
Building rule-based unique identifiers from structured HR data


# EXCEL_TASK-6

## 📅Membership Duration & Pass/Fail Status (Date & Logical Functions)

A solved Excel practice file with two exercises covering date calculations and conditional logic with formatting.

🔧 What's inside

Question 1 — Membership Duration Tracker
Calculates how long each customer has been a member, from their signup date to today, in both months and days:

excel
=DATEDIF(D5,TODAY(),"M")   → Membership Duration (Months)
=DATEDIF(D5,TODAY(),"D")   → Membership Duration (Days)

Question 2 — Pass/Fail Evaluator with Conditional Formatting
Evaluates student marks and flags each as Pass/Fail, then visually highlights the result:

excel
=IF(D7>=40,"PASS","FAIL")
----------------------------------------------------------------------------------------------------------------------------------------------------------------
✅ "PASS" → Green fill
----------------------------------------------------------------------------------------------------------------------------------------------------------------
❌ "FAIL" → Red fill
(via Conditional Formatting rules on the Status column)
----------------------------------------------------------------------------------------------------------------------------------------------------------------
🧠 Skills demonstrated
DATEDIF() for calculating elapsed time between two dates (months & days)
TODAY() for dynamic, self-updating date calculations
IF() for conditional logic / pass-fail classification
Conditional Formatting to visually encode categorical results



	
