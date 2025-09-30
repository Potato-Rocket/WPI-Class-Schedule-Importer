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

# discards any uscheduled sections, then any empty courses or time frames
print()
print("Verifying schedule data...")
# for each time frame (iterating over key list copy)
for time_frame in list(time_frames.keys()):
    # get the list of courses in the time frame
    courses_ = time_frames[time_frame]

    # iterate over a copy of the course list
    for course in list(courses_):
        # get the list of sections in the course
        sections_ = courses[course]

        # iterate over a copy of the sections list
        for section in list(sections_):

            # check whether the section has a schedule
            if section['Meeting Patterns'] is None:
                # display a message
                code = " ".join(section['Section'].split()[:2])
                print(f"Discarded section {code} because not scheduled.")
                sections_.remove(section)  # remove it from the course
                sections.remove(section)  # remove it from the main list

        # check whether the course has any section left
        if len(sections_) == 0:
            # display a message
            code = " ".join(course.split()[:2])
            print(f"Discarded course {code} due to no scheduled sections.")
            courses_.remove(course)  # remove it from the time frame
            del courses[course]  # remove it from the main dict

    # check whether the time frame has any courses left
    if len(courses_) == 0:
        # display a message
        time = f"{datetime.strftime(time_frame[0], r"%Y-%m-%d")} to {datetime.strftime(time_frame[1], r"%Y-%m-%d")}"
        print(f"Discarded time frame from {time} due to no remaining courses.")
        del time_frames[time_frame]  # remove from the dict

# displays a time frame -> course -> section tree view with schedule info
print()
print(f"{len(sections)} sections, {len(courses.keys())} courses, and {len(time_frames.keys())} time frames found!")

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
