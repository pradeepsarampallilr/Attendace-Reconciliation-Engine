# Attendance Reconciliation Engine

This project reads a raw access-control movement report, cleans it, matches check-ins and check-outs, optionally enriches the result with a shift roster, and exports an Excel workbook with:

- an Attendance sheet
- an Audit Log sheet

It is designed for monthly door-access data where each row contains employee movement information and repeated swipe columns such as "In" and "Out".

---

## Current Processing Rules

- The current-month movement report is required. A previous-month movement report is optional.
- When a previous-month report is supplied, only its final transaction date is loaded. Those events are placed into the employee swipe pools before the current-month events so a previous-month check-in can match a checkout recorded on the first day of the current month.
- Previous-month events are used only for matching. Visits whose check-in belongs to the previous month are excluded from the current-month Attendance output, so their hours are not included in current-month totals.
- Each check-in is paired with the earliest unused checkout that occurs after it, before the next check-in, on the same day or the following day, and within the configured maximum shift duration of 18 hours.
- A missing checkout remains blank and is recorded in the Audit Log. An unused checkout is treated as an orphan checkout and is also recorded in the Audit Log.
- For overnight shifts, a post-midnight check-in before the roster shift end belongs to the previous work date. For example, `2026-08-04 00:52` belongs to the `2026-08-03` shift for a `4:00 PM - 1:30 AM` roster schedule.
- If a current report ends with a check-in that has no checkout and the employee has an overnight roster shift, the engine uses the next day’s roster shift-end time as the fallback checkout. For example, an August 31 check-in for a `4:30 PM - 1:30 AM` shift receives a fallback checkout of September 1 at `1:30 AM`.
- The `Date` and `Daily Total Hours` columns are grouped by employee and work date. The daily hours are summed across all visits, written on the first row of the group, and visually merged across the group in Excel. Original visit rows are retained.

---

## What the engine does

The engine performs a pipeline of work:

1. Reads the movement report from CSV or Excel
2. Finds the real header row even if the exported file contains title rows or blank rows at the top
3. Normalizes employee IDs so numbers stored as strings, floats, or padded values all match consistently
4. Extracts employee-specific swipe lists for check-ins and check-outs
5. Reconciles each check-in to a valid checkout by applying business rules
6. Optionally attaches roster metadata such as employee name, email, and shift timing
7. Writes the final attendance rows and audit log to an Excel workbook

---

## Inputs

### 1) Movement report (required)
This is the raw swipe report exported from the access-control system.

Typical fields include:

- employee code / employee id
- employee name
- transdate / date
- columns labeled "In" and "Out" with time values

The code accepts both:

- .csv files
- .xlsx files

### 2) Shift roster (optional)
This is an optional Excel file containing employee metadata such as:

- employee id
- legal name
- email
- shift timing

If provided, the engine uses it to enrich attendance rows with roster details.

---

## Output

The script writes a workbook with two sheets:

### Attendance
Contains one row per check-in and includes columns like:

- Employee Name
- Employee id
- Email
- Shift Timings
- Year
- Month
- Week
- Checked In Date
- Checked Out Date
- No.of hours in Office

### Audit Log
Contains rows that were skipped, dropped, or left unresolved, with details like:

- employee id
- employee name
- issue type
- details
- timestamp

---

## Main flow of the engine

The flow is implemented in the `AttendanceEngine` class.

### 1) `run()`
This is the public entry point.

```python
def run(self) -> tuple[list[VisitRecord], list[AuditEntry]]:
```

It does the following:

- loads the shift roster if a shift file was supplied
- loads the movement report
- extracts swipe events by employee
- reconciles each employee’s check-ins and check-outs
- enriches each visit with roster details
- sorts visits by employee name and check-in time
- returns both the visit list and the audit entries

---

## Step-by-step breakdown

### 2) `_load_movement_table()`
This is the file-loading phase.

It does the following:

- reads the CSV or Excel file into a Pandas DataFrame
- stores the source filename in `raw.attrs["source"]`
- detects the real header row
- stores the header row as the column names
- discards any rows before the official header
- returns the header list and the cleaned data rows

This is important because raw exports often include titles, notes, or blank rows before the actual table begins.

---

### 3) `_detect_header_row()`
This function scans the first few rows to locate the real header row.

It looks for a column named `empcode` (case-insensitive), and once found, it uses that row as the official header.

If it cannot find the required header pattern, it raises a clear error.

This prevents the engine from accidentally treating a title row as the real data header.

---

### 4) `_extract_events()`
This is the raw-data cleaning and grouping phase.

The function:

- lowercases the headers
- locates the important columns such as employee code, name, date, and in/out swipe fields
- reads each row one at a time with `iterrows()`
- normalizes the employee id
- parses dates and times
- combines date + time into a full `datetime`
- groups all check-ins and check-outs by employee in a dictionary

The output is shaped like this:

```python
{
    "101": ("Alice", [datetime(...), ...], [datetime(...), ...]),
    "102": ("Bob", [...], [...]),
}
```

This is the structure needed before pairing the swipes.

---

### 5) `_reconcile_employee()`
This is the pairing logic.

For each employee, it iterates through all check-ins in order and tries to find a valid corresponding checkout.

Important behavior:

- each checkout is used only once
- a checkout is not allowed to be used after the next check-in
- a checkout must not be too far away from the check-in
- a checkout must usually be on the same day or the very next day
- a check-in with no valid checkout remains in the attendance output with a blank checkout
- a leftover checkout with no valid check-in is recorded as an orphan and logged in the Audit Log

This function decides whether a swipe becomes a valid attendance record or an audit issue.

---

### 6) `_find_checkout()`
This is the validation rule engine for pairing a single check-in with a checkout.

It checks multiple constraints:

- the candidate checkout must be after the check-in
- if there is a next check-in, the checkout cannot be after that next check-in
- the checkout must be on either the same date or the next date
- the duration between check-in and checkout must not exceed the configured shift window

If all checks pass, it returns the checkout timestamp; otherwise it returns `None`.

---

## Dataclasses used in the engine

### `ShiftRecord`
```python
@dataclass(frozen=True, slots=True)
class ShiftRecord:
    emp_id: str
    name: str
    email: str
    shift_timing: str
```

This is a roster record. It is intended to be a stable lookup object. The values are meant to remain fixed once loaded.

Why `frozen=True`:

- prevents accidental mutation of roster values
- treats the record as a fixed, trusted data object

Why `slots=True`:

- uses a compact internal storage layout for memory efficiency
- common for simple data containers

---

### `VisitRecord`
```python
@dataclass(slots=True)
class VisitRecord:
    emp_code: str
    name: str
    check_in: datetime
    check_out: Optional[datetime] = None
    email: str = ""
    shift_timing: str = ""
```

This represents one attendance visit and stores the raw data that gets written into the final output.

#### Properties on `VisitRecord`

##### `year`
Returns the year from `check_in`.

```python
@property
def year(self) -> int:
    return self.check_in.year
```

##### `month_name`
Formats the month as a full month name.

```python
@property
def month_name(self) -> str:
    return self.check_in.strftime("%B")
```

##### `week_of_month`
Calculates the week number of the month using the date.

```python
@property
def week_of_month(self) -> int:
    return ((self.check_in.day - 1) // 7) + 1
```

##### `hours_in_office`
Computes the attendance duration in hours.

```python
@property
def hours_in_office(self) -> Optional[float]:
    if self.check_out is None:
        return None
    delta = self.check_out - self.check_in
    return round(delta.total_seconds() / 3600, 2)
```

The column `# of hours in Office` is calculated this way.

##### `as_row()`
Converts the object into a list of output values in the exact column order required for Excel export.

---

### `AuditEntry`
```python
@dataclass(slots=True)
class AuditEntry:
    emp_code: str
    name: str
    issue: str
    detail: str
    timestamp: Optional[datetime] = None
```

This stores issues found during reconciliation, such as:

- orphan checkout
- no matching checkout
- employee not found in shift roster
- date parsing issue

It is later exported to the Audit Log sheet.

---

## Helper functions and why they matter

### `normalize_id(value)`
This cleans employee IDs into one canonical format.

Example transformations:

- "158920" -> "158920"
- 158920.0 -> "158920"
- " 158920 " -> "158920"

This is necessary because Excel and CSV exports often represent the same employee id differently across files.

---

### `parse_date(value)`
This converts mixed Excel/CSV date values into Python `date` objects.

It handles:

- `None`
- `NaN`
- pandas timestamps
- Python `date` values
- strings like "2026-09-01"

---

### `parse_time(value)`
This converts time-like values into Python `time` objects.

It handles strings such as:

- "08:30"
- "08:30:00"

and also handles pandas / datetime values and missing values.

---

### `clean_text(value)`
This strips whitespace and converts common missing values such as `NaN` to empty strings.

---

## Important configuration values

### `HEADER_PROBE = "empcode"`
This tells the loader what field to search for when identifying the real header row.

### `MAX_HEADER_SCAN_ROWS = 15`
The loader checks only the first 15 rows for the header, which is enough for most exported files but still prevents infinite scanning.

### `MAX_SHIFT_HOURS = 18`
This is the maximum allowed duration between a check-in and checkout for a single visit.

This prevents unrealistic matches when swipe data is incomplete or missing.

---

## How the final report is written

The `ReportWriter.write()` method does the final export.

It:

- converts each `VisitRecord` to a row using `as_row()`
- converts each `AuditEntry` to a row using `as_row()`
- loads them into pandas DataFrames
- writes both to an Excel workbook with two sheets
- styles headers and date columns

The workbook is written using `openpyxl`.

---

## How to run the engine

From the project folder:

```bash
python3 attendance_engine.py
```

The script will ask for:

1. path to the movement report (.xlsx or .csv)
2. path to the shift roster file (optional)
3. output file name (optional)

Then it generates the Excel report automatically.

---

## Common project customization ideas

### 1) Modify the shift window to 24 hours

The shift window is controlled by:

```python
MAX_SHIFT_HOURS = 18
```

If you want to allow longer overnight shifts, change it to 24:

```python
MAX_SHIFT_HOURS = 24
```

This change affects the rules in `_find_checkout()`, which rejects checkout values whose gap exceeds this threshold.

This is useful when:

- shift rosters include long overnight work
- employees sometimes work late into the next morning
- the access system records shifts beyond 18 hours

> Note: this is a business rule and should be adjusted to match the real attendance policy.

---

### 2) Fill orphan check-ins without a checkout

The script currently does this when a check-in has no valid checkout:

- it creates a `VisitRecord` with `check_out = None`
- it logs an audit entry saying "No matching checkout"
- it leaves the checkout blank in the attendance sheet

This is intentionally conservative and avoids inventing data.
If you want to fill missing checkouts automatically, the usual approach is to add a rule in `_reconcile_employee()` or `_find_checkout()` such as:
- use a fixed default checkout time for the same day if no valid exit is found
- use the next midnight or a configured shift-end time when the business rules allow it
- apply a fallback only under a clearly defined policy

Examples of fallback policies:
- same-day checkout = shift end value from the roster
- missing checkout = next midnight if the employee is still inside a plausible overnight shift window
- leave it blank unless a business rule explicitly says to estimate it

This project chooses the safer option by default: leave the value blank rather than guessing.

---

## Design notes

This code is intentionally conservative. It tries to avoid wrong matches and false attendance data.

That is why it:
- rejects orphan checkouts
- refuses to bridge across too long a gap
- refuses matches that cross into the next check-in period
- leaves blanks instead of guessing

This makes it safer for payroll, audit, and operations reporting.

---

## Summary

The engine is a data cleaning and reconciliation pipeline:
- raw movement report
- detect and normalize headers
- parse and clean employee swipe data
- match check-ins and check-outs under business rules
- enrich with roster metadata
- export a workbook and audit log

The logic is intentionally strict so it prefers valid data over guessed data.
---

## Notes

This README documents the current logic and where to change behavior if required.

The main places to tune this project are:

- `AttendanceEngine.MAX_SHIFT_HOURS` for broader or narrower shift windows
- `_reconcile_employee()` for pairing strategy
- `_find_checkout()` for validation rules
- `_enrich_with_roster()` for roster enrichment behavior

No code changes are required to understand or use the engine as described here.
