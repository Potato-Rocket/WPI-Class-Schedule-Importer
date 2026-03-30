# Class Schedule Importer

A Python TUI tool that converts university class schedules from Excel spreadsheets into iCalendar files for easy import into calendar applications like Outlook.

## Overview

This tool helps you automatically add your class schedule to your calendar by:
1. Reading your exported Workday schedule (Excel format)
2. Parsing class meeting times, locations, and details
3. Allowing you to select which sections to export
4. Generating a calendar file (.ics) you can import into Outlook or other calendar apps

## Prerequisites

- Python 3.x
- Required Python packages:
  ```bash
  pip install -r requirements.txt
  ```

## Getting Your Class Schedule from Workday

1. Log in to **Workday**
2. Navigate to **Academics** → **View My Courses**
3. Click the **Export to Excel** button (located in the top right corner of the courses table)
4. Save the downloaded Excel file to a location you can easily find

## Running the Script

1. Run the script:
   ```bash
   python3 class_schedule.py
   ```

2. Follow the prompts:
   - Select your first Excel file when the file dialog appears
   - **Optional**: Load additional files (e.g., Fall and Spring semesters together)
     - The script will prompt: "Load another file? (y/n)"
     - All files must have the same column structure
   - The script will analyze your schedule and group courses by time frame
   - Interactively select which time frame, courses, and sections you want to export
   - Choose where to save your calendar file (.ics)

## Adding the Calendar to Outlook

### Outlook Desktop (Windows/Mac)

1. Open **Outlook**
2. Go to **File** → **Open & Export** → **Import/Export**
3. Select **Import an iCalendar (.ics) or vCalendar file (.vcs)**
4. Click **Next**
5. Browse to your exported `.ics` file
6. Choose **Import** (adds to your existing calendar) or **Open as New** (creates a separate calendar)
7. Click **OK**

### Outlook Web (outlook.office.com)

1. Log in to **Outlook Web**
2. Click the **Calendar** icon (bottom left)
3. Click **Add calendar** (left sidebar)
4. Select **Upload from file**
5. Click **Browse** and select your `.ics` file
6. Choose which calendar to import into
7. Click **Import**

### Outlook Mobile (iOS/Android)

1. Email the `.ics` file to yourself
2. Open the email on your mobile device
3. Tap the `.ics` attachment
4. Tap **Add to Calendar** or **Import**
5. Select the calendar to add events to

## Known Limitations

- Events are created as individual instances (not recurring series)
- Some calendar applications may have formatting issues
- Timezone is set to America/New_York

## Troubleshooting

**Script won't run:**
- Ensure Python 3 is installed: `python3 --version`
- Install dependencies: `pip install openpyxl icalendar`

**Excel file won't load:**
- Make sure you exported directly from Workday using the "Export to Excel" button
- Verify the file is in `.xlsx` format

**"Error: First header cell is empty!" or "Error: Missing required headers!":**
- Your Excel file may not match the expected Workday format
- The script expects headers at row 3, column 2 (after title row and label column)
- If your file has a different structure, adjust `ROW_SKIP` and `COL_SKIP` constants at the top of `class_schedule.py`
- Required columns: Course Listing, Section, Meeting Patterns, Start Date, End Date

**"Error: Headers don't match previous file!":**
- When loading multiple files, all must have the same column structure
- This file will be skipped - you can continue with already loaded files
- Ensure all files are Workday exports from the same system

**"Warning: Reached MAX limit":**
- Your spreadsheet has more than 100 rows
- Increase the `MAX` constant at the top of `class_schedule.py` (e.g., change to `MAX = 200`)

**Events not showing in calendar:**
- Check that you imported the file (not just opened it)
- Verify the date range matches your current semester
- Ensure your calendar app supports `.ics` files

## License

MIT License - See LICENSE file for details.
