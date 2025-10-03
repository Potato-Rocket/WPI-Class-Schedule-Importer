import tkinter as tk
from tkinter import filedialog
import openpyxl
import warnings
import sys
import uuid
from datetime import datetime, timedelta, timezone
from icalendar import Calendar, Event, vDatetime, vRecur

MAX = 100  # a reasonable limit on searching for data cells
COL_SKIP = 1  # number of columns to skip before data starts
ROW_SKIP = 2  # number of rows to skip before data starts

WEEKDAYS = {
    'M': "MO",
    'T': "TU",
    'W': "WE",
    'R': "TH",
    'F': "FR"
}


def get_filename():
    """
    Open a system dialog to allow the user to select a .xlsx file.
    Returns the filename as a string.
    """
    print("\nPlease select an excel file to read.")

    # start up tkinter for the file dialog
    root = tk.Tk()
    root.withdraw()

    # acquire the input file with OS file dialog
    fname = filedialog.askopenfilename(filetypes=[('Microsoft Excel Files', '*.xlsx')])

    # close tkinter now that we're done with it
    root.destroy()
    return fname


def parse_spreadsheet(fname):
    """
    Parses the given .xlsx file and extracts section information.
    Returns a list of sections, where each section is a dictionary of its attributes with the spreadsheet header as the key.    
    """
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    # try to open the file
    try:
        print("Loading excel spreadsheet...")
        wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
        ws = wb.active
    except openpyxl.utils.exceptions.InvalidFileException:
        # exits if dialog was cancelled
        print("Invalid file!")
        sys.exit()
    
    start_row = ROW_SKIP + 1
    start_col = COL_SKIP + 1

    print("Finding headers...")
    # parses the spreadsheet to find column names
    headers = []
    for col in range(MAX):
        # traverses the third row, starting on the second column
        cell = ws.cell(start_row, col + start_col).value
        # only adds the cell if the column exists
        if cell is not None:
            headers.append(cell)
    
    start_row += 1  # moves down to the next row for section data

    print("Finding sections...")
    # parses section info from each row
    sections = []
    for row in range(MAX):

        if ws.cell(row + start_row, start_col).value is None:
            break
        
        section = {'Index': row}
        for col in range(len(headers)):
            section[headers[col]] = ws.cell(row + start_row, col + start_col).value
        
        if section['Meeting Patterns'] is not None:
            fields = section['Meeting Patterns'].split(" | ")

            time_str = fields[1].split(" - ")
            section['Start Time'] = datetime.strptime(time_str[0], r"%H:%M")
            section['End Time'] = datetime.strptime(time_str[1], r"%H:%M")

            section['Location'] = fields[2]
            section['Meeting Patterns'] = fields[0].split("-")
        
        sections.append(section)

    return sections


def group_data(sections):
    """
    Groups sections into courses, and courses into time frames.
    Returns two dictionaries: courses and time_frames.
    """
    print("Grouping sections into courses...")
    # groups sections into courses
    courses = {}
    for section in sections:
        course = section['Course Listing']  # identifies the course name
        if course in courses.keys():
            courses[course].append(section)  # appends to entry if not new
        else:
            courses[course] = [section]  # makes new entry if new

    print("Grouping courses into time frames...")
    # groups courses into time frames
    time_frames = {}
    for course, course_sections in courses.items():
        # a time frame is tuple of start and end date
        time_frame = (course_sections[0]['Start Date'], course_sections[0]['End Date'])
        if time_frame in time_frames:  # appends to entry if not new
            time_frames[time_frame].append(course)
        else:  # makes new entry if new
            time_frames[time_frame] = [course]
    
    return courses, time_frames


def verify_scheduling(sections, courses, time_frames):
    """
    Verifies that each section has a schedule, removing unscheduled sections, courses with no scheduled sections, and time frames with no courses. Modifies the input lists and dictionaries in place.
    """
    print("\nVerifying schedule data...")
    # for each time frame (iterating over key list copy)
    for time_frame in list(time_frames.keys()):
        # get the list of courses in the time frame
        tf_courses = time_frames[time_frame]

        # iterate over a copy of the course list
        for course in list(tf_courses):
            # get the list of sections in the course
            course_sections = courses[course]

            # iterate over a copy of the sections list
            for section in list(course_sections):

                # check whether the section has a schedule
                if section['Meeting Patterns'] is None:
                    # display a message
                    code = " ".join(section['Section'].split()[:2])
                    print(f"Discarded section {code} because not scheduled.")
                    course_sections.remove(section)  # remove it from the course
                    sections.remove(section)  # remove it from the main list

            # check whether the course has any section left
            if len(course_sections) == 0:
                # display a message
                code = " ".join(course.split()[:2])
                print(f"Discarded course {code} due to no scheduled sections.")
                tf_courses.remove(course)  # remove it from the time frame
                del courses[course]  # remove it from the main dict

        # check whether the time frame has any courses left
        if len(tf_courses) == 0:
            # display a message
            time_str = f"{datetime.strftime(time_frame[0], r"%Y-%m-%d")} to {datetime.strftime(time_frame[1], r"%Y-%m-%d")}"
            print(f"Discarded time frame from {time_str} due to no remaining courses.")
            del time_frames[time_frame]  # remove from the dict


def print_data_summary(sections, courses, time_frames):
    """Prints a summary of the length of data found."""
    # displays information about the courses
    print(f"\n{len(sections)} sections, {len(courses.keys())} courses, and {len(time_frames.keys())} time frames found!")


def print_tree_view(courses, time_frame):
    """Prints a tree view of the time frames, courses, and sections."""

    print()
    print("=" * 60)
    print("Tree view of time frames, courses, and sections")
    print("=" * 60)

    # for each time frame
    for time_frame, tf_courses in time_frames.items():
        # print the time frame
        print(f"\n{datetime.strftime(time_frame[0], r"%Y-%m-%d")} to {datetime.strftime(time_frame[1], r"%Y-%m-%d")}:")

        # for each course in the time frame
        for i, course in enumerate(tf_courses):
            # print the course name and gutter tree
            print("│")
            last = i == len(tf_courses) - 1
            char = '└' if last else '├'
            print(f"{char}── {course}")

            # identify the sections
            course_sections = courses[course]

            # for each section in the course
            for section in course_sections:
                # print section name and gutter tree
                char = '└' if section['Index'] == course_sections[-1]['Index'] else '├'
                tree = ' ' if last else '│'
                code = " ".join(section['Section'].split()[:2])
                print(f"{tree}   {char}── {code:<14}", end="")

                # print meeting patterns
                out = ""
                if section['Meeting Patterns'] is not None:
                    out = ", ".join(section['Meeting Patterns'])
                print(f"    {out:<13}", end="")

                # print start time
                out = "--:--"
                if 'Start Time' in section.keys():
                    out = section['Start Time'].strftime(r"%H:%M")
                print(f"    {out}", end="")

                # print end time
                out = "--:--"
                if 'End Time' in section.keys():
                    out = section['End Time'].strftime(r"%H:%M")
                print(f" - {out}")


def select_sections(time_frames, courses):
    """
    Interactively select which sections to include in calendar export.
    Returns a list of approved sections.
    """
    approved_sections = []

    print()
    print("=" * 60)
    print("Select sections to export to calendar")
    print("=" * 60)

    # iterate over each time frame
    for time_frame, tf_courses in time_frames.items():
        time_str = f"{datetime.strftime(time_frame[0], r'%Y-%m-%d')} to {datetime.strftime(time_frame[1], r'%Y-%m-%d')}"
        print(f"\nTime frame: {time_str}")

        while True:
            choice = input("  (y)es / (n)o / (s)pecific courses? ").lower().strip()
            if choice in ['y', 'n', 's']:
                break
            print("Invalid input. Please enter y, n, or s.")

        if choice == 'y':
            # approve all sections in this time frame
            for course in tf_courses:
                approved_sections.extend(courses[course])
            print(f"  ✓ Added all courses in this time frame")

        elif choice == 's':
            # drill down to course selection
            for course in tf_courses:
                print(f"\n    Course: {course}")

                while True:
                    choice = input("      (y)es / (n)o / (s)pecific courses? ").lower().strip()
                    if choice in ['y', 'n', 's']:
                        break
                    print("    Invalid input. Please enter y, n, or s.")

                if choice == 'y':
                    # approve all sections in this course
                    approved_sections.extend(courses[course])
                    print(f"      ✓ Added all sections in this course")

                elif choice == 's':
                    # drill down to section selection
                    course_sections = courses[course]
                    for section in course_sections:
                        code = " ".join(section['Section'].split()[:2])

                        # show schedule info
                        days = ", ".join(section['Meeting Patterns'])
                        start = section['Start Time'].strftime(r"%H:%M")
                        end = section['End Time'].strftime(r"%H:%M")

                        print(f"\n        Section: {code} ({days} {start}-{end})")

                        while True:
                            choice = input("          (y)es / (n)o? ").lower().strip()
                            if choice in ['y', 'n']:
                                break
                            print("        Invalid input. Please enter y or n.")

                        if choice == 'y':
                            approved_sections.append(section)
                            print(f"          ✓ Added section")
                        else:
                            print(f"          ✗ Skipped section")

                else:
                    # choice == 'n' means skip this course
                    print(f"      ✗ Skipped course")

        else:
            # choice == 'n' means skip this time frame
            print(f"  ✗ Skipped time frame")
    
    return approved_sections


def generate_calendar(sections):
    """
    Generates an iCalendar file from the given sections.
    """
    print("\nGenerating iCalendar data...")
    cal = Calendar()
    cal.add('prodid', '-//Class Schedule Importer//mxm.dk//')
    cal.add('version', '2.0')

    for section in sections:
        event = Event()
        event.add('uid', uuid.uuid4().hex.upper())
        event.add('summary', section['Section'])
        event.add('location', section['Location'])

        tz = timezone(timedelta(hours=-4))
        dtstart = datetime.combine(section['Start Date'].date(), section['Start Time'].time(), tzinfo=tz)
        dtend = datetime.combine(section['Start Date'].date(), section['End Time'].time(), tzinfo=tz)
        event.add('dtstart', dtstart)
        event.add('dtend', dtend)

        description = f"{section['Delivery Mode']} {section['Instructional Format']} with Professor {section['Instructor']}.\n\nGenerated by Class Schedule Importer."
        event.add('description', description)

        cal.add_component(event)

    return cal


def save_calendar(cal):
    """
    Saves the given iCalendar object to a .ics file.
    """
    print("Please select a location to save the iCalendar file.")
    # start up tkinter for the file dialog
    root = tk.Tk()
    root.withdraw()

    # acquire the output file with OS file dialog
    fname = f"class_schedule_{datetime.now().strftime(r'%Y%m%d')}.ics"
    fname = filedialog.asksaveasfilename(defaultextension=".ics", initialfile=fname, filetypes=[('iCalendar Files', '*.ics')])

    root.destroy()

    # exits if dialog was cancelled
    if fname == "":
        print("Save cancelled.")
        return
    # ensure the filename ends with .ics
    if not fname.endswith(".ics"):
        fname += ".ics"
    
    # write the calendar to the file
    with open(fname, 'wb') as f:
        f.write(cal.to_ical())
    print(f"iCalendar data saved to {fname}")


# main script execution
print("Welcome to Class Schedule Importer!")
fname = get_filename()
sections = parse_spreadsheet(fname)
courses, time_frames = group_data(sections)
print_data_summary(sections, courses, time_frames)
verify_scheduling(sections, courses, time_frames)
print_data_summary(sections, courses, time_frames)
print_tree_view(courses, time_frames)
approved_sections = select_sections(time_frames, courses)
print(f"\n{len(approved_sections)} sections out of {len(sections)} approved for export.")
calendar = generate_calendar(approved_sections)
save_calendar(calendar)
print("\nThank you for using Class Schedule Importer!")
