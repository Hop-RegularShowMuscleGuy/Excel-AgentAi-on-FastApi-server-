# Excel Functions Knowledge Base
# Each section covers one function or pattern.
# Use in apply_formula tool: formula_template with {row} as row number placeholder.

---

## IF
Checks a condition and returns one value if true, another if false.

**Syntax:** `=IF(condition, value_if_true, value_if_false)`

**Examples with {row}:**
- `=IF(B{row}<>"", "FILLED", "EMPTY")` — checks if cell B is not empty
- `=IF(A{row}>0, "POSITIVE", "NEGATIVE")` — checks sign of a number
- `=IF(C{row}="YES", 1, 0)` — converts text to 0/1
- `=IF(D{row}>=100, "OK", "BELOW TARGET")` — compares against a threshold
- `=IF(AND(A{row}>0, B{row}>0), "BOTH POSITIVE", "NOT BOTH")` — compound condition

**Nested IF:**
- `=IF(A{row}>90, "A", IF(A{row}>70, "B", IF(A{row}>50, "C", "FAIL")))` — grade classification

---

## IFS
Checks multiple conditions in order, returns value for the first one that is true.

**Syntax:** `=IFS(condition1, value1, condition2, value2, ...)`

**Examples with {row}:**
- `=IFS(A{row}>90,"A",A{row}>70,"B",A{row}>50,"C",TRUE,"FAIL")` — grades without nesting
- `=IFS(B{row}="PL","POLAND",B{row}="DE","GERMANY",B{row}="FR","FRANCE",TRUE,"OTHER")`

---

## VLOOKUP
Searches for a value in the first column of a range and returns a value from another column in the same row.

**Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
- `range_lookup`: 0 or FALSE = exact match, 1 or TRUE = approximate match

**Examples with {row}:**
- `=VLOOKUP(A{row}, Sheet2!A:B, 2, 0)` — looks up A{row} in Sheet2, returns column 2
- `=VLOOKUP(B{row}, $D$1:$F$100, 3, 0)` — fixed table in the same sheet, returns column 3
- `=IFERROR(VLOOKUP(A{row}, Sheet2!A:C, 2, 0), "NOT FOUND")` — with error handling
- `=VLOOKUP(A{row}&B{row}, Sheet2!A:D, 3, 0)` — combined key lookup

---

## XLOOKUP
Modern replacement for VLOOKUP — searches in any direction, no column index needed.

**Syntax:** `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])`

**Examples with {row}:**
- `=XLOOKUP(A{row}, Sheet2!A:A, Sheet2!B:B, "NOT FOUND")` — simple lookup with fallback
- `=XLOOKUP(A{row}, Sheet2!A:A, Sheet2!C:C, 0)` — returns column C
- `=XLOOKUP(A{row}&B{row}, Sheet2!A:A&Sheet2!B:B, Sheet2!C:C, "MISSING")` — compound key

---

## INDEX + MATCH
Flexible alternative to VLOOKUP — can search in any column.

**Syntax:** `=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))`

**Examples with {row}:**
- `=INDEX(Sheet2!B:B, MATCH(A{row}, Sheet2!A:A, 0))` — equivalent to VLOOKUP
- `=IFERROR(INDEX(Sheet2!C:C, MATCH(A{row}, Sheet2!A:A, 0)), "N/A")` — with error handling
- `=INDEX(Sheet2!D:D, MATCH(1, (Sheet2!A:A=A{row})*(Sheet2!B:B=B{row}), 0))` — multi-condition match (array formula, Ctrl+Shift+Enter)

---

## SUMIF
Sums cells that meet a single condition.

**Syntax:** `=SUMIF(range, criteria, [sum_range])`

**Examples with {row}:**
- `=SUMIF(B:B, A{row}, C:C)` — sum of column C where column B equals A{row}
- `=SUMIF(A:A, ">"&D{row}, B:B)` — sum of B where A is greater than D{row}
- `=SUMIF(C:C, C{row}, D:D)` — sum for all rows sharing the same category as current row

---

## SUMIFS
Sums cells that meet multiple conditions.

**Syntax:** `=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)`

**Examples with {row}:**
- `=SUMIFS(D:D, B:B, B{row}, C:C, C{row})` — sum of D for the same combination of B and C
- `=SUMIFS(E:E, A:A, A{row}, D:D, "ACTIVE")` — sum of E where A matches and status is ACTIVE
- `=SUMIFS(C:C, A:A, ">="&DATE(2024,1,1), A:A, "<="&DATE(2024,12,31))` — sum for a full year

---

## COUNTIF
Counts cells that meet a single condition.

**Syntax:** `=COUNTIF(range, criteria)`

**Examples with {row}:**
- `=COUNTIF(A:A, A{row})` — how many times the value in A{row} appears in column A
- `=COUNTIF(B:B, ">"&B{row})` — how many values in B are greater than B{row}
- `=COUNTIF($A$1:A{row}, A{row})` — occurrence number of the current value (expanding range)

---

## COUNTIFS
Counts cells that meet multiple conditions.

**Syntax:** `=COUNTIFS(criteria_range1, criteria1, criteria_range2, criteria2, ...)`

**Examples with {row}:**
- `=COUNTIFS(A:A, A{row}, B:B, B{row})` — detects duplicates based on two columns
- `=COUNTIFS(C:C, C{row}, D:D, "YES")` — count of records with the same category and status YES

---

## TEXT
Converts a value to text in a specified format.

**Syntax:** `=TEXT(value, format_text)`

**Examples with {row}:**
- `=TEXT(A{row}, "YYYY-MM-DD")` — date as ISO string
- `=TEXT(B{row}, "0.00")` — number with 2 decimal places
- `=TEXT(C{row}, "#,##0")` — number with thousands separator
- `=TEXT(A{row}, "DD/MM/YYYY")` — date in European format
- `=TEXT(B{row}, "0%")` — value as percentage

---

## CONCATENATE / & / CONCAT / TEXTJOIN
Combines text strings.

**Syntax:**
- `=A{row}&" "&B{row}` — & operator (simplest)
- `=CONCAT(A{row}, " ", B{row})` — function form
- `=TEXTJOIN(", ", TRUE, A{row}, B{row}, C{row})` — with separator, ignores empty cells

**Examples with {row}:**
- `=A{row}&" "&B{row}` — first name and last name
- `=UPPER(A{row})&"_"&TEXT(B{row},"000")` — key with zero-padding
- `=TEXTJOIN("; ", TRUE, A{row}, B{row}, C{row})` — columns joined with semicolon

---

## LEFT / RIGHT / MID
Extracts substrings from text.

**Syntax:**
- `=LEFT(text, num_chars)` — characters from the left
- `=RIGHT(text, num_chars)` — characters from the right
- `=MID(text, start_num, num_chars)` — characters from the middle

**Examples with {row}:**
- `=LEFT(A{row}, 3)` — first 3 characters
- `=RIGHT(A{row}, 4)` — last 4 characters
- `=MID(A{row}, 4, 2)` — 2 characters starting at position 4
- `=LEFT(A{row}, FIND(" ", A{row})-1)` — part before the first space (first name)

---

## LEN
Returns the length of a text string.

**Syntax:** `=LEN(text)`

**Examples with {row}:**
- `=LEN(A{row})` — character count of A{row}
- `=IF(LEN(A{row})>10, "LONG", "SHORT")` — classify by length

---

## TRIM
Removes extra spaces from text — leading, trailing, and internal duplicates.

**Syntax:** `=TRIM(text)`

**Examples with {row}:**
- `=TRIM(A{row})` — clean up whitespace
- `=TRIM(UPPER(A{row}))` — clean and convert to uppercase

---

## UPPER / LOWER / PROPER
Changes text casing.

**Syntax:** `=UPPER(text)` / `=LOWER(text)` / `=PROPER(text)`

**Examples with {row}:**
- `=UPPER(A{row})` — ALL CAPS
- `=LOWER(A{row})` — all lowercase
- `=PROPER(A{row})` — First Letter Of Each Word Capitalized

---

## SUBSTITUTE
Replaces occurrences of a substring within a text string.

**Syntax:** `=SUBSTITUTE(text, old_text, new_text, [instance_num])`

**Examples with {row}:**
- `=SUBSTITUTE(A{row}, " ", "_")` — spaces to underscores
- `=SUBSTITUTE(A{row}, ",", ".")` — commas to dots (decimal numbers)
- `=SUBSTITUTE(SUBSTITUTE(A{row}, " ", ""), "-", "")` — remove spaces and hyphens

---

## FIND / SEARCH
Returns the position of a substring within a text string.
- `FIND` — case-sensitive
- `SEARCH` — case-insensitive

**Syntax:** `=FIND(find_text, within_text, [start_num])`

**Examples with {row}:**
- `=FIND("@", A{row})` — position of @ character (for email)
- `=IFERROR(FIND("@", A{row}), 0)` — returns 0 if not found
- `=LEFT(A{row}, FIND("@", A{row})-1)` — part before the @ sign

---

## IFERROR
Catches errors and returns a fallback value instead of #N/A, #DIV/0!, etc.

**Syntax:** `=IFERROR(value, value_if_error)`

**Examples with {row}:**
- `=IFERROR(VLOOKUP(A{row}, Sheet2!A:B, 2, 0), "NOT FOUND")` — handle #N/A
- `=IFERROR(B{row}/C{row}, 0)` — handle division by zero
- `=IFERROR(VALUE(A{row}), 0)` — handle text-to-number conversion error

---

## VALUE
Converts a text string representing a number into an actual number.

**Syntax:** `=VALUE(text)`

**Examples with {row}:**
- `=VALUE(A{row})` — converts text "123" to number 123
- `=IFERROR(VALUE(A{row}), 0)` — safe conversion

---

## DATE
Creates a date from year, month, and day values.

**Syntax:** `=DATE(year, month, day)`

**Examples with {row}:**
- `=DATE(A{row}, B{row}, C{row})` — date from three separate columns
- `=DATE(YEAR(A{row}), MONTH(A{row})+1, 1)-1` — last day of the month from A{row}

---

## YEAR / MONTH / DAY
Extracts parts of a date.

**Examples with {row}:**
- `=YEAR(A{row})` — year
- `=MONTH(A{row})` — month number (1–12)
- `=DAY(A{row})` — day of month (1–31)
- `=TEXT(A{row}, "YYYY-MM")` — year and month as text

---

## DATEDIF
Calculates the difference between two dates in a specified unit.

**Syntax:** `=DATEDIF(start_date, end_date, unit)`
- Units: "Y" years, "M" months, "D" days, "YM" months ignoring years, "MD" days ignoring months

**Examples with {row}:**
- `=DATEDIF(A{row}, TODAY(), "Y")` — age in years
- `=DATEDIF(A{row}, B{row}, "D")` — number of days between two dates
- `=DATEDIF(A{row}, TODAY(), "M")` — how many months have passed

---

## TODAY / NOW
Returns the current date or current date and time.

**Examples with {row}:**
- `=TODAY()-A{row}` — number of days since the date in A{row}
- `=IF(A{row}<TODAY(), "OVERDUE", "OK")` — checks if a deadline has passed
- `=DATEDIF(A{row}, TODAY(), "Y")` — age calculation

---

## ROUND / ROUNDUP / ROUNDDOWN / MROUND
Rounds a number to a specified number of digits.

**Syntax:** `=ROUND(number, num_digits)`

**Examples with {row}:**
- `=ROUND(A{row}, 2)` — 2 decimal places
- `=ROUND(A{row}/B{row}*100, 1)` — percentage with 1 decimal
- `=ROUNDUP(A{row}, 0)` — always round up to whole number
- `=ROUNDDOWN(A{row}, 0)` — always round down to whole number
- `=MROUND(A{row}, 5)` — round to nearest multiple of 5

---

## SUM / MAX / MIN / AVERAGE
Basic row-level aggregation.

**Examples with {row}:**
- `=SUM(B{row}:F{row})` — sum of columns B through F in the current row
- `=MAX(B{row}:F{row})` — maximum value in the row
- `=MIN(B{row}:F{row})` — minimum value in the row
- `=AVERAGE(B{row}:F{row})` — average value in the row

---

## AND / OR / NOT
Logical operators for combining conditions.

**Examples with {row}:**
- `=AND(A{row}>0, B{row}>0)` — TRUE if both are positive
- `=OR(A{row}="YES", B{row}="YES")` — TRUE if either is YES
- `=NOT(A{row}="")` — TRUE if A{row} is not empty
- `=IF(AND(A{row}<>"", B{row}<>""), "COMPLETE", "INCOMPLETE")` — both columns filled

---

## RANK
Returns the rank of a number within a list.

**Syntax:** `=RANK(number, ref, [order])`
- `order`: 0 = descending (default), 1 = ascending

**Examples with {row}:**
- `=RANK(B{row}, B:B, 0)` — rank from highest to lowest
- `=RANK(B{row}, $B$2:$B$1000, 1)` — rank from lowest to highest (fixed range)

---

## UNIQUE (dynamic arrays)
Returns unique values from a range. Available in Excel 365 / Excel 2021+.

**Note:** Use without {row} — this is a single-cell spill formula.
- `=UNIQUE(A2:A100)` — list of unique values from column A

---

## ROW
Returns the row number of a cell.

**Examples with {row}:**
- `=ROW()` — current row number
- `=ROW()-1` — sequential record number (when row 1 is the header)

---

## CHOOSE
Returns a value from a list based on an index number.

**Syntax:** `=CHOOSE(index_num, value1, value2, ...)`

**Examples with {row}:**
- `=CHOOSE(WEEKDAY(A{row}), "Sun","Mon","Tue","Wed","Thu","Fri","Sat")` — day of week name
- `=CHOOSE(A{row}, "LOW", "MEDIUM", "HIGH")` — maps 1/2/3 to labels

---

## NETWORKDAYS
Calculates the number of working days between two dates.

**Syntax:** `=NETWORKDAYS(start_date, end_date, [holidays])`

**Examples with {row}:**
- `=NETWORKDAYS(A{row}, B{row})` — working days between A and B
- `=NETWORKDAYS(A{row}, TODAY())` — working days from a date until today

---

## ISBLANK
Checks whether a cell is empty.

**Examples with {row}:**
- `=ISBLANK(A{row})` — TRUE if A is empty
- `=IF(ISBLANK(A{row}), "MISSING", A{row})` — fallback for empty cells

---

## ISNUMBER
Checks whether a value is a number.

**Examples with {row}:**
- `=ISNUMBER(A{row})` — TRUE if the value is a number
- `=IF(ISNUMBER(A{row}), A{row}*1.2, 0)` — apply multiplier only to numbers

---

## ISTEXT
Checks whether a value is text.

**Examples with {row}:**
- `=ISTEXT(A{row})` — TRUE if the value is text

---

## EXACT
Compares two text strings — case-sensitive.

**Syntax:** `=EXACT(text1, text2)`

**Examples with {row}:**
- `=EXACT(A{row}, "admin")` — TRUE only for "admin", not "Admin"
- `=IF(EXACT(A{row}, B{row}), "MATCH", "NO MATCH")`

---

## Coalesce pattern (chained IFERROR + VLOOKUP)
Fallback through multiple lookup tables.

**Example with {row}:**
```
=IFERROR(
  VLOOKUP(A{row}, Sheet2!A:B, 2, 0),
  IFERROR(
    VLOOKUP(A{row}, Sheet3!A:B, 2, 0),
    "NOT FOUND"
  )
)
```

---

## Composite key pattern
Combines multiple columns into a single lookup key.

**Examples with {row}:**
- `=VLOOKUP(A{row}&"|"&B{row}, Sheet2!A:C, 3, 0)` — requires a helper column in Sheet2 with the same separator
- `=XLOOKUP(A{row}&B{row}, Sheet2!A:A&Sheet2!B:B, Sheet2!C:C, "NOT FOUND")` — no helper column needed

---

## Running total pattern
Cumulative sum with an expanding range.

**Examples with {row}:**
- `=SUM($B$2:B{row})` — sum from B2 to the current row
- `=COUNTIF($A$2:A{row}, A{row})` — occurrence number of the current value

---

## Dense rank pattern
Rank without gaps (no skipped positions for ties).

**Example with {row}:**
- `=SUMPRODUCT((B$2:B$1000>B{row})/COUNTIF(B$2:B$1000, B$2:B$1000))+1` — dense rank descending