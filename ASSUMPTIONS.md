# Spreadsheet Format Assumptions & Fragility Analysis

This document outlines all assumptions made about the Workday Excel export format and identifies fragile points in the parsing logic.

## Overview

The script is specifically designed for **WPI Workday exports** and makes several assumptions about spreadsheet structure. Some are safe for this specific use case, others are potential failure points.

---

## ✅ SAFE ASSUMPTIONS (WPI Workday Specific)

These assumptions are safe because they're controlled by WPI's Workday export format:

### Column Names
- **Assumption**: Headers match exactly: `Course Listing`, `Section`, `Meeting Patterns`, `Start Date`, `End Date`, `Instructor`, `Delivery Mode`, `Instructional Format`
- **Location**: Lines 109, 120, 150, 158, 206, 321-322, 330
- **Risk**: LOW - Workday column names are standardized
- **Impact if wrong**: KeyError crash when accessing dictionary keys

### Meeting Patterns Format
- **Assumption**: `Meeting Patterns` cell follows format: `"M-T-W-R-F | HH:MM - HH:MM | Location"`
- **Location**: Lines 85-93
- **Current parsing**:
  ```python
  fields = section['Meeting Patterns'].split(" | ")  # splits on pipe
  time_str = fields[1].split(" - ")                   # expects time in fields[1]
  section['Start Time'] = datetime.strptime(time_str[0], r"%H:%M")
  section['End Time'] = datetime.strptime(time_str[1], r"%H:%M")
  section['Location'] = fields[2]                     # location in fields[2]
  section['Meeting Patterns'] = fields[0].split("-")  # days split by dash
  ```
- **Risk**: LOW for Workday, but format-dependent
- **Impact if wrong**: IndexError or ValueError crash

### Day Codes
- **Assumption**: Days use single letters: M, T, W, R, F (not 'H' for Thursday)
- **Location**: Lines 14-20 (WEEKDAYS constant)
- **Risk**: LOW - Workday uses this standard
- **Impact if wrong**: Days won't map to iCalendar format correctly

### Time Format
- **Assumption**: Times are in 24-hour format `HH:MM` (e.g., "14:00" not "2:00 PM")
- **Location**: Line 89
- **Risk**: LOW - Workday exports use 24-hour format
- **Impact if wrong**: ValueError during datetime parsing

---

## ⚠️ FRAGILE POINTS (Easy to Break)

These assumptions could fail with minor variations:

### 1. **Fixed Row/Column Offset**
- **Assumption**: Headers always start at row 3, column 2 (after 2 header rows, 1 label column)
- **Location**: Lines 11-12, 58-59
  ```python
  COL_SKIP = 1  # skip 1 column
  ROW_SKIP = 2  # skip 2 rows
  start_row = ROW_SKIP + 1  # = 3
  start_col = COL_SKIP + 1  # = 2
  ```
- **Risk**: MEDIUM - Workday might change export format
- **Impact if wrong**: Reads wrong cells, gets garbage data or crashes
- **Fix difficulty**: Easy - make these configurable constants

### 2. **Hard Limit on Data Size**
- **Assumption**: No more than 100 rows or columns of data
- **Location**: Lines 10, 64, 76
  ```python
  MAX = 100
  for col in range(MAX):  # header search
  for row in range(MAX):  # section search
  ```
- **Risk**: MEDIUM - Students with many courses or wide spreadsheets fail silently
- **Impact if wrong**: Truncates data without warning
- **Note**: `ws.max_row` and `ws.max_column` both return 1 with `read_only=True, data_only=True`, so manual iteration is necessary
- **Fix difficulty**: Easy - add warning when approaching limit, or make MAX configurable via CLI args

### 3. **Naive Empty Cell Detection**
- **Assumption**: First empty cell in column A indicates end of data
- **Location**: Lines 78-79
  ```python
  if ws.cell(row + start_row, start_col).value is None:
      break
  ```
- **Risk**: LOW-MEDIUM - Could stop early if a cell is blank
- **Impact if wrong**: Misses rows after first blank
- **Fix difficulty**: Medium - need more robust end-of-data detection

### 4. **Meeting Patterns Split Logic**
- **Assumption**: Days are always separated by dashes, never spaces or other delimiters
- **Location**: Line 93
  ```python
  section['Meeting Patterns'] = fields[0].split("-")  # expects "M-T-R-F"
  ```
- **Risk**: MEDIUM - If Workday changes to "M T R F" or "MTRF", this breaks
- **Impact if wrong**: Gets entire string as single day, or crashes iCalendar generation
- **Fix difficulty**: Easy - add alternative parsing logic

### 5. **Date Type Assumptions**
- **Assumption**: `Start Date` and `End Date` cells contain datetime objects (not strings)
- **Location**: Lines 120, 325-326
- **Risk**: MEDIUM - openpyxl `data_only=True` returns calculated values, should be datetime
- **Impact if wrong**: AttributeError when calling `.date()` method
- **Fix difficulty**: Medium - add type checking and string parsing fallback

### 6. **Timezone Hardcoding**
- **Assumption**: All events are in UTC-4 (Eastern Daylight Time)
- **Location**: Line 324
  ```python
  tz = timezone(timedelta(hours=-4))
  ```
- **Risk**: MEDIUM - Wrong during winter (should be UTC-5 EST) or for other timezones
- **Impact if wrong**: Events appear at wrong times
- **Fix difficulty**: Medium - need proper timezone library (pytz) or DST logic

---

## 🔥 CRITICAL ASSUMPTIONS (Will Cause Crashes)

These assumptions will cause immediate failures if violated:

### 1. **Non-Null Meeting Patterns for Scheduled Sections**
- **Assumption**: Scheduled sections have non-None `Meeting Patterns`, unscheduled sections have None
- **Location**: Lines 85, 148, 211
- **Risk**: HIGH if assumption is wrong, LOW if Workday is consistent
- **Impact if wrong**:
  - Line 85: Tries to call `.split()` on None → AttributeError
  - Unscheduled sections not properly filtered
- **Fix difficulty**: Already handled by verify_scheduling(), but parse_spreadsheet() crashes first

### 2. **Section Dictionary Keys Exist**
- **Assumption**: Selected sections always have these keys: `Meeting Patterns`, `Start Time`, `End Time`, `Start Date`, `End Date`, `Location`, `Section`, `Delivery Mode`, `Instructional Format`, `Instructor`
- **Location**: Throughout selection and calendar generation (lines 280-282, 321-331)
- **Risk**: MEDIUM - If verify_scheduling() fails or is bypassed
- **Impact if wrong**: KeyError crash during selection UI or calendar generation
- **Fix difficulty**: Easy - add .get() with defaults

### 3. **Index Field for Sorting**
- **Assumption**: All sections have an `Index` field for identifying last item
- **Location**: Lines 81, 204
- **Risk**: LOW - Always set during parsing
- **Impact if wrong**: Comparison fails, tree view breaks
- **Fix difficulty**: Easy - use enumerate() instead

---

## 💡 REINFORCEMENT RECOMMENDATIONS

### Quick Wins (Easy to Fix)

1. **Make offsets configurable**:
   ```python
   # At top of file, with documentation
   COL_SKIP = 1  # Workday exports: skip 1 label column
   ROW_SKIP = 2  # Workday exports: skip 2 header rows (title + blank)
   ```

2. **~~Use worksheet dimensions instead of MAX~~**:
   - **NOTE**: `ws.max_row` and `ws.max_column` both return 1 with `read_only=True, data_only=True`
   - Manual iteration with MAX limit is currently necessary
   - Could add warning if hitting MAX limit

3. **Add defensive dictionary access**:
   ```python
   # Instead of: section['Location']
   section.get('Location', 'TBD')
   ```

4. **Validate Meeting Patterns format**:
   ```python
   if section['Meeting Patterns'] and " | " in section['Meeting Patterns']:
       fields = section['Meeting Patterns'].split(" | ")
       if len(fields) >= 3:  # validate expected structure
           # ... existing parsing
   ```

### Medium Effort

5. **Add date type checking**:
   ```python
   start_date = section['Start Date']
   if isinstance(start_date, str):
       start_date = datetime.strptime(start_date, "%Y-%m-%d")
   ```

6. **Better timezone handling**:
   ```python
   import pytz
   tz = pytz.timezone('America/New_York')  # handles DST automatically
   ```

### Nice to Have

7. **Command-line argument support**:
   ```bash
   python3 class_schedule.py --input Fall_Semester.xlsx --output fall.ics
   python3 class_schedule.py --row-skip 2 --col-skip 1
   ```
   - Useful for automation and batch processing
   - Can override defaults without GUI

8. **Validation warnings**:
   ```python
   if ws.max_row > 100:
       print(f"Warning: Spreadsheet has {ws.max_row} rows, only processing first 100")
   ```

---

## Current State Assessment

### What's Solid
- ✅ Data grouping logic (time frames → courses → sections)
- ✅ Interactive selection UI with drill-down
- ✅ Validation and discarding of unscheduled items
- ✅ Error handling for file selection cancellation

### What's Fragile
- ⚠️ Hardcoded offsets and limits
- ⚠️ String parsing without validation
- ⚠️ Timezone handling
- ⚠️ Assumption that Workday format never changes

### What's Broken (Known Issues)
- ❌ iCalendar export format (Excel can't open)
- ❌ No recurrence rules (single events only)
- ❌ Wrong start dates (semester start vs. first class day)

---

## Intentional Decisions for WPI Workday

Since this tool is specifically for WPI Workday exports, the following assumptions are **intentionally kept**:

1. **Column names** - Safe to hardcode Workday's exact column names
2. **Meeting Patterns format** - Workday consistently uses `"Days | Time | Location"`
3. **Day codes** - M/T/W/R/F is Workday standard
4. **24-hour time** - Workday exports use this format

These don't need defensive coding unless Workday changes their export format.
