# Excel Lookup Functions: Zero to Hero

A comprehensive hands-on guide covering Excel's core lookup functions: `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `INDEX+MATCH`, and `XLOOKUP`, complete with syntax, practical examples, and limitations.

---

## 1. Sample Dataset

Use this base table for the vertical lookup examples below:

<img width="401" height="182" alt="image" src="https://github.com/user-attachments/assets/a6a9f56f-a34f-4ed0-8eb1-53662c490b04" />

---

## 2. VLOOKUP (Vertical Lookup)

Searches for a value in the first column of a table and returns a value in the same row from a specified column.

* **Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
* **Limitation:** Cannot look to its left (lookup value must strictly be in the leftmost column).

### Example
Find the **Department** of Employee `E103`:

<img width="539" height="288" alt="7" src="https://github.com/user-attachments/assets/59d13bef-1a02-4279-b7b3-45dd1af4df21" />

Result: Finance

Line-by-line Explanation:

"E103": The value we want to search for.

A2:D5: The entire data table range.

3: Department is the 3rd column from EmpID.

FALSE: Specifies an exact match.

---

## 3. HLOOKUP (Horizontal Lookup)

Searches for a value in the top row of a table and returns a value in the same column from a specified row.

### Sample Horizontal Dataset
<img width="454" height="132" alt="image" src="https://github.com/user-attachments/assets/d1c62838-4f41-4c70-97ed-786c294071da" />

* **Syntax:** `=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

### Example

Find the **Margin** for `Q3`:

<img width="541" height="209" alt="8" src="https://github.com/user-attachments/assets/dfcd5305-2053-477b-aeed-1b295f12718b" />

* **Result:** `25%`
* **Line-by-line Explanation:**
* `"Q3"`: The quarter we are searching across Row 1.
* `B1:E3`: The data grid containing metrics across quarters.
* `2`: Margin is in the 2nd row of this selected range.
* `FALSE`: Exact match.



---

## 4. MATCH

Returns the relative numeric position of an item in a single row or column.

* **Syntax:** `=MATCH(lookup_value, lookup_array, [match_type])`

### Example

Find the row position of `Amit` in the Name column (`B2:B5`):

<img width="474" height="261" alt="9" src="https://github.com/user-attachments/assets/a9ba81e1-ba79-4948-b76f-ee3a28608d4a" />


* **Result:** `3`
* **Line-by-line Explanation:**
* `"Amit"`: Target value.
* `B2:B5`: The single column array containing names.
* `0`: Exact match mode (returns 3 because Amit is the 3rd entry in B2:B5).



---

## 5. INDEX

Returns the value of a cell at a given row and column coordinate within a range.

* **Syntax:** `=INDEX(array, row_num, [column_num])`

### Example

Extract the value at Row 4, Column 4 from `A2:D5`:

<img width="445" height="270" alt="10" src="https://github.com/user-attachments/assets/5c195aab-2467-440b-abf1-1dab9c07da8b" />


* **Result:** `90000`
* **Line-by-line Explanation:**
* `A2:D5`: Lookup grid.
* `4`: 4th row of the selection (Sneha's record).
* `4`: 4th column of the selection (Salary column).



---

## 6. INDEX + MATCH (Dynamic Dual Lookup)

Combines INDEX and MATCH to overcome VLOOKUP limitations. It can look left, insert columns without breaking formulas, and perform dynamic 2-way lookups.

### Example (Left Lookup)

Given `Salary = 72000`, find the corresponding **Name**:

<img width="588" height="268" alt="11" src="https://github.com/user-attachments/assets/ffde7bee-d362-4070-b9ba-c9b2e0fa187e" />


* **Result:** `Amit`
* **Line-by-line Explanation:**
* `MATCH(72000, D2:D5, 0)` executes first: scans the Salary column (`D2:D5`) and finds 72000 at relative position **3**.
* `INDEX(B2:B5, 3)` executes second: retrieves the 3rd item from the Name column (`B2:B5`), returning **Amit**.

---

## 7. XLOOKUP (Modern Standard)

Replaces VLOOKUP, HLOOKUP, and INDEX+MATCH. Works in any direction (left/right/up/down), defaults to exact match, and has built-in error handling.

* **Syntax:** `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`

### Example 1: Basic Left Lookup with Error Handling

Find the **EmpID** for Name `Priya`:

<img width="623" height="274" alt="12" src="https://github.com/user-attachments/assets/a53177fa-9423-450b-b17d-0091d69b0d49" />

* **Result:** `E102`
* **Line-by-line Explanation:**
* `"Priya"`: Item to find.
* `B2:B5`: Lookup column (where to look).
* `A2:A5`: Return column (what to fetch — works leftward without issue).
* `"Not Found"`: Custom fallback value if no record matches (eliminates the need for `IFERROR`).



### Example 2: Two-Way Matrix Lookup (Dynamic Row & Column)

Formula:

<img width="727" height="268" alt="13" src="https://github.com/user-attachments/assets/d2dcada5-f672-40b6-9cc6-44f53c8460db" />


* **Result:** `90000`
* **Explanation:** The inner XLOOKUP dynamically picks the column based on the header "Salary", and the outer XLOOKUP filters by row for "Sneha".

---

## 8. Summary Comparison

| Feature | VLOOKUP | HLOOKUP | INDEX + MATCH | XLOOKUP |
| --- | --- | --- | --- | --- |
| **Direction** | Vertical (Right only) | Horizontal (Down only) | Any (Left/Right/Up/Down) | Any |
| **Default Match** | Approximate (TRUE) | Approximate (TRUE) | Requires `0` parameter | Exact (`0`) |
| **Column Insert Safe?** | No (hardcoded index breaks) | No | Yes (dynamic arrays) | Yes |
| **Built-in Error Catch** | No (needs `IFERROR`) | No | No | Yes (`if_not_found`) |
| **Excel Version** | All | All | All | Excel 2021 / Microsoft 365 |

