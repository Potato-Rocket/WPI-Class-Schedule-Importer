# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Class Schedule Importer is a Python GUI tool that converts university class schedules from Excel spreadsheets (.xlsx) into iCalendar (.ics) files for import into calendar applications.

### Core Workflow

1. User selects an Excel file via GUI dialog (tkinter)
2. Spreadsheet is parsed to extract section data with headers starting at row 3, column 2
3. Sections are grouped into courses, and courses into time frames (start/end date pairs)
4. Unscheduled sections/courses are validated and discarded
5. User interactively selects which sections to export (time frame → course → section drill-down)
6. Selected sections are converted to iCalendar events
7. Calendar file is saved via GUI dialog

## Running the Application

```bash
python3 class_schedule.py
```

The script runs as a standalone executable from start to finish. No command-line arguments needed - all input is through GUI dialogs and interactive prompts.

## Architecture

### Data Processing Pipeline

**parse_spreadsheet()** → **group_data()** → **verify_scheduling()** → **select_sections()** → **generate_calendar()** → **save_calendar()**

### Data Structure

Sections are dictionaries with these key fields after parsing:
- Original spreadsheet columns (Course Listing, Section, Instructor, etc.)
- Parsed from 'Meeting Patterns' string: `Start Time`, `End Time`, `Location`, `Meeting Patterns` (now a list of day codes)
- `Index`: Row number for tracking

### Key Constants

- `MAX = 100`: Search limit for finding data cells in spreadsheet
- `COL_SKIP = 1`, `ROW_SKIP = 2`: Offset to data start (headers at row 3, col 2)
- `WEEKDAYS`: Maps single-letter day codes (M/T/W/R/F) to iCalendar format (MO/TU/WE/TH/FR)

### Known Issues

Calendar export has multiple issues (see commit 77dfe85):
- **Generated .ics files cannot be opened by Excel** - format incompatibility
- **Recurrence logic is not implemented** - events are created as single instances, not recurring series
- **Start dates are incorrect** - events use the semester start date instead of calculating the first occurrence based on meeting patterns (e.g., a Tuesday/Thursday class starting on a semester that begins Monday will incorrectly start on Monday instead of Tuesday)
- Timezone is hardcoded to UTC-4 (line 324)

## Dependencies

- `openpyxl`: Excel file reading
- `icalendar`: iCalendar generation
- `tkinter`: GUI file dialogs (built-in to Python)

Install with:
```bash
pip install openpyxl icalendar
```
