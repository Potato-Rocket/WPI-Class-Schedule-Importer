import tkinter as tk
from tkinter import filedialog
import openpyxl
import warnings
import sys
from datetime import datetime

MAX = 100  # a reasonable limit on searching for data cells

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

print("Please select a file.")

# start up tkinter for the file dialog
root = tk.Tk()
root.withdraw()

# acquire the input file with OS file dialog
fname = filedialog.askopenfilename(filetypes=[('Microsoft Excel Files', '*.xlsx')])

# close tkinter now that we're done with it
root.destroy()

# display the selected filename
print("File selected:")
print(fname)

# try to open the file
# exits if dialog was cancelled
try:
    wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
    ws = wb.active
except openpyxl.utils.exceptions.InvalidFileException:
    print("Invalid file!")
    sys.exit()

print("\nParsing spreadsheet data...")

# parses the spreadsheet to find column names
headers = []
for col in range(MAX):
    # traverses the third row, starting on the second column
    cell = ws.cell(3, col + 2).value
    # only adds the cell if the column exists
    if cell is not None:
        headers.append(cell)

# parses section info from each row
sections = []
for row in range(MAX):

    if ws.cell(row + 4, 2).value is None:
        break

    section = {'Index': row}
    for col in range(len(headers)):
        section[headers[col]] = ws.cell(row + 4, col + 2).value
    
    if section['Meeting Patterns'] is not None:
        fields = section['Meeting Patterns'].split(" | ")

        time = fields[1].split(" - ")
        section['Start Time'] = datetime.strptime(time[0], r"%H:%M")
        section['End Time'] = datetime.strptime(time[1], r"%H:%M")

        section['Location'] = fields[2]
        section['Meeting Patterns'] = fields[0].split("-")
    
    sections.append(section)

# groups sections into courses
courses = {}
for section in sections:
    course = section['Course Listing']  # identifies the course name
    if course in courses.keys():
        courses[course].append(section)  # appends to entry if not new
    else:
        courses[course] = [section]  # makes new entry if new

# groups courses into time frames
time_frames = {}
for course, sections_ in courses.items():
    # a time frame is tuple of start and end date
    time_frame = (sections_[0]['Start Date'], sections_[0]['End Date'])
    if time_frame in time_frames:  # appends to entry if not new
        time_frames[time_frame].append(course)
    else:  # makes new entry if new
        time_frames[time_frame] = [course]

# displays information about the courses
print(f"{len(sections)} sections, {len(courses.keys())} courses, and {len(time_frames.keys())} time frames found!")

# displays a time frame -> course -> section tree view with schedule info

# for each time frame
for time_frame, courses_ in time_frames.items():
    # print the time frame
    print()
    print(f"{datetime.strftime(time_frame[0], r"%Y-%m-%d")} to {datetime.strftime(time_frame[1], r"%Y-%m-%d")}:")

    # for each course in the time frame
    for i, course in enumerate(courses_):
        # print the course name and gutter tree
        print("│")
        last = i == len(courses_) - 1
        char = '└' if last else '├'
        print(f"{char}── {course}")

        # identify the sections
        sections_ = courses[course]

        # for each section in the course
        for section in sections_:
            # print section name and gutter tree
            char = '└' if section['Index'] == sections_[-1]['Index'] else '├'
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
