# Class Schedule Importer

A Python GUI tool that converts university class schedules from Excel spreadsheets into iCalendar files for easy import into calendar applications like Outlook.

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
  pip install openpyxl icalendar
  ```

## Getting Your Class Schedule from Workday

1. Log in to **Workday**
2. Navigate to **Academics** → **View My Courses**
3. Click the **Export to Excel** button (located in the top right corner of the courses table)
4. Save the downloaded Excel file to a location you can easily find

## Running the Script

1. Make sure you've installed the required dependencies:
   ```bash
   pip install openpyxl icalendar
   ```

2. Run the script:
   ```bash
   python3 class_schedule.py
   ```

3. Follow the prompts:
   - Select your exported Excel file when the file dialog appears
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

## How It Works

The script processes your Excel file through the following pipeline:

1. **Parse Spreadsheet**: Extracts section data starting from row 3, column 2
2. **Group Data**: Organizes sections into courses and time frames (semester start/end dates)
3. **Validate Scheduling**: Removes any unscheduled classes or sections
4. **Interactive Selection**: Lets you choose which sections to export
5. **Generate Calendar**: Converts selected sections to iCalendar events
6. **Save Calendar**: Exports as a `.ics` file

## Data Structure

Each class section includes:
- Course information (name, listing, section number)
- Instructor details
- Meeting times and days
- Location
- Start and end dates

## Known Limitations

- Timezone is currently hardcoded to UTC-4
- Recurrence logic for repeating events is under development
- Currently creates individual event instances rather than recurring series

## Troubleshooting

**Script won't run:**
- Ensure Python 3 is installed: `python3 --version`
- Install dependencies: `pip install openpyxl icalendar`

**Excel file won't load:**
- Make sure you exported directly from Workday using the "Export to Excel" button
- Verify the file is in `.xlsx` format

**Events not showing in calendar:**
- Check that you imported the file (not just opened it)
- Verify the date range matches your current semester
- Ensure your calendar app supports `.ics` files

## Dependencies

- `openpyxl`: Excel file reading
- `icalendar`: iCalendar file generation
- `tkinter`: GUI file dialogs (included with Python)

## License

This is a personal utility script. Use and modify as needed.
