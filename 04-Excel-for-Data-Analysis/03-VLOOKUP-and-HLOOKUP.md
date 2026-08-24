# VLOOKUP and HLOOKUP in Excel

VLOOKUP and HLOOKUP are useful Excel functions for finding information from a table. They help reduce manual searching and make data analysis faster and more efficient.

---

## 1. VLOOKUP – Vertical Lookup

**VLOOKUP** searches for a value vertically in the first column of a table and returns a related value from another column in the same row.

### Syntax

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
```

### Example

Suppose we have the following employee data:

| Employee ID | Name   | Department     |
| ----------- | ------ | -------------- |
| 101         | Shreya | Data Analytics |
| 102         | Rahul  | HR             |
| 103         | Neha   | Finance        |
| 104         | Aman   | Marketing      |

If we want to find the name of Employee ID **103**:

```excel
=VLOOKUP(103,A2:C5,2,FALSE)
```

**Result:** `Neha`

### How VLOOKUP works

* **103** → Value we want to find
* **A2:C5** → Range containing the data
* **2** → Column number from which we want the result
* **FALSE** → Exact match

### Key Points

* VLOOKUP searches **vertically**.
* The lookup value must be in the **first column** of the selected table.
* `FALSE` is used when an exact match is required.
* It returns the value from the specified column in the same row.

---

## 2. HLOOKUP – Horizontal Lookup

**HLOOKUP** searches for a value horizontally in the first row of a table and returns a related value from another row in the same column.

### Syntax

```excel
=HLOOKUP(lookup_value, table_array, row_index_num, FALSE)
```

### Example

Suppose we have product data arranged horizontally:

| Product ID | P101 | P102 | P103 | P104 |
| ---------- | ---: | ---: | ---: | ---: |
| Price      |  500 |  800 | 1200 |  950 |

If we want to find the price of **P103**:

```excel
=HLOOKUP("P103",A1:E2,2,FALSE)
```

**Result:** `1200`

### How HLOOKUP works

* **P103** → Value we want to find
* **A1:E2** → Range containing the data
* **2** → Row number from which we want the result
* **FALSE** → Exact match

### Key Points

* HLOOKUP searches **horizontally**.
* The lookup value must be in the **first row** of the selected table.
* `FALSE` is used for an exact match.
* It returns the value from the specified row in the same column.

---

## 3. VLOOKUP vs HLOOKUP

| Feature          | VLOOKUP                     | HLOOKUP                       |
| ---------------- | --------------------------- | ----------------------------- |
| Full Form        | Vertical Lookup             | Horizontal Lookup             |
| Search Direction | Vertical                    | Horizontal                    |
| Searches In      | First column                | First row                     |
| Returns From     | Another column              | Another row                   |
| Best Used When   | Data is arranged vertically | Data is arranged horizontally |

### Easy Way to Remember

**VLOOKUP → V = Vertical ↓**

**HLOOKUP → H = Horizontal →**

---

## 4. Why Use Lookup Functions?

Lookup functions are useful when working with large datasets because they can:

* Reduce manual searching
* Retrieve information quickly
* Connect related data
* Improve productivity
* Reduce repetitive work
* Make reports easier to prepare

---

## 5. Exact Match Using FALSE

For most basic lookup tasks, using `FALSE` is helpful when we want an exact result.

Example:

```excel
=VLOOKUP(103,A2:C5,2,FALSE)
```

Here, Excel looks specifically for **103** instead of finding an approximate value.

---

## 6. Common Lookup Errors

### `#N/A`

The lookup value could not be found in the selected range.

### `#REF!`

The formula contains an invalid cell or column/row reference.

### `#VALUE!`

The formula contains an incorrect value or argument.

---

## 7. Practical Use in Data Analysis

VLOOKUP and HLOOKUP can be used in different situations, such as:

* Finding employee information
* Retrieving product prices
* Matching IDs with names
* Combining information from different tables
* Preparing reports
* Working with large Excel datasets

---

## Key Takeaway

**VLOOKUP and HLOOKUP make it easier to find and retrieve information from structured data.**

Understanding these functions is an important step toward becoming more efficient with Excel and building strong **Data Analytics skills**.

**Learn → Practice → Apply → Improve 📊**
