# Excel Lookup Functions: Zero to Hero

A comprehensive hands-on guide covering Excel's core lookup functions: `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `INDEX+MATCH`, and `XLOOKUP`, complete with syntax, practical examples, and limitations.

---

## 1. Sample Dataset

Use this base table for the vertical lookup examples below:

| EmpID | Name    | Department  | Salary  |
|:------|:--------|:------------|:--------|
| E101  | Rahul   | Analytics   | 65000   |
| E102  | Priya   | Marketing   | 58000   |
| E103  | Amit    | Finance     | 72000   |
| E104  | Sneha   | Engineering | 90000   |

---

## 2. VLOOKUP (Vertical Lookup)

Searches for a value in the first column of a table and returns a value in the same row from a specified column.

* **Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
* **Limitation:** Cannot look to its left (lookup value must strictly be in the leftmost column).

### Example
Find the **Department** of Employee `E103`:

=VLOOKUP("E103", A2:D5, 3, FALSE)

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

| Metric | Q1 | Q2 | Q3 | Q4 |
| --- | --- | --- | --- | --- |
| Sales | 12000 | 15000 | 18000 | 22000 |
| Margin | 20% | 22% | 25% | 28% |

* **Syntax:** `=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

### Example

Find the **Margin** for `Q3`:

```excel
=HLOOKUP("Q3", B1:E3, 2, FALSE)

```

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

```excel
=MATCH("Amit", B2:B5, 0)

```

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

```excel
=INDEX(A2:D5, 4, 4)

```

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

```excel
=INDEX(B2:B5, MATCH(72000, D2:D5, 0))

```

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

```excel
=XLOOKUP("Priya", B2:B5, A2:A5, "Not Found")

```

* **Result:** `E102`
* **Line-by-line Explanation:**
* `"Priya"`: Item to find.
* `B2:B5`: Lookup column (where to look).
* `A2:A5`: Return column (what to fetch — works leftward without issue).
* `"Not Found"`: Custom fallback value if no record matches (eliminates the need for `IFERROR`).



### Example 2: Two-Way Matrix Lookup (Dynamic Row & Column)

Formula:

```excel
=XLOOKUP("Sneha", B2:B5, XLOOKUP("Salary", A1:D1, A2:D5))

```

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

