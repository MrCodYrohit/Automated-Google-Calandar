from __future__ import print_function

import csv
from datetime import date, datetime, timedelta, time
import re
import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import sys

from datetime import datetime, timedelta
import os.path
#!pip install tkcalendar
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import csv
from datetime import datetime


days = {
    "mon": "monday",
    "tues": "tuesday",
    "wed": "wednesday",
    "thur": "thursday",
    "fri": "friday"
}

month_captial_written = {
    "january": "01",
    "february": "02",
    "march": "03",
    "april": "04",
    "may": "05",
    "june": "06",
    "july": "07",
    "august": "08",
    "september": "09",
    "october": "10",
    "november": "11",
    "december": "12"
}
month_small_written = {
    "jan": "01",
    "feb": "02",
    "mar": "03",
    "apr": "04",
    "may": "05",
    "june": "06",
    "jul": "07",
    "aug": "08",
    "sep": "09",
    "oct": "10",
    "nov": "11",
    "dec": "12"
}

# 1 file cover

# for academic calender
check_month = [
    "january", "february", "march", "april", "may", "june", "july", "august",
    "september", "october", "november", "december", "jan", "feb", "mar", "apr",
    "may", "june", "jul", "aug", "sep", "oct", "nov", "dec"
]
check_days = [
    "mon", "monday", "tues", "tuesday", "Wed", "wednesday", "thur", "thursday",
    "fri", "friday", "sunday", "sun", "sat", "saturday"
]
#

# 2 file cover
# for slot and first
# "Monday","Tuesday","Wednesday",,"Thursday","Friday","Sunday","Saturday"

# 1file
# for the slotdata
capital_days = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"]
#

day_number = {
    "MONDAY": 0,
    "TUESDAY": 1,
    "WEDNESDAY": 2,
    "THURSDAY": 3,
    "FRIDAY": 4
}
short_term = {
    "MONDAY": "MO",
    "TUESDAY": "TU",
    "WEDNESDAY": "WE",
    "THURSDAY": "TH",
    "FRIDAY": "FR"
}

year = datetime.now().year
slot = {}
main_list = []
first_year_data = {}
slot_data = []
not_in_slot = []

subsitute = {}
holiday = {}
all = []

sub = subsitute
hol = holiday

#year = 2023     #Enter current year


def compare_dates(date_str1, date_str2):
  """
  Compare two date strings in the format "DD/MM/YYYY".

  Parameters:
  - date_str1 (str): The first date string.
  - date_str2 (str): The second date string.

  Returns:
  - int: Returns 1 if date_str2 is after date_str1, 0 if they are equal, and raises a ValueError if the date strings are     
   invalid.
  """
  format_str = "%d/%m/%Y"

  try:
    date1 = datetime.strptime(date_str1, format_str)
    date2 = datetime.strptime(date_str2, format_str)

    if date1 < date2:
      return 1
    elif date1 > date2:
      return 0
    else:
      return 0

  except ValueError as e:
    return f"Error: {e}. Please provide valid date strings in the format DD/MM/YYYY."


def correct_order(s_date, e_date, mid_date, holy_date):
  """
  Check the chronological order of dates in a semester-related sequence.

  Parameters:
  - s_date (str): The starting date of the semester.
  - e_date (str): The ending date of the semester.
  - mid_date (str): The date of the midsem break.
  - holy_date (str): The holy date.

  Returns:
  - int: Returns 0 if the dates are in the correct order,
          1 if the last date of the midsem break is wrong,
          2 if the starting date of the midsem exam is wrong,
          3 if the ending date of the semester is wrong.

  Note: Assumes that the compare_dates function is available to compare date strings.
  """

  if (compare_dates(s_date, e_date)):
    if (compare_dates(s_date, mid_date) and compare_dates(mid_date, e_date)):
      if (compare_dates(mid_date, holy_date)
          and compare_dates(holy_date, e_date)):
        return 0  #Every date is fine
      else:
        return 1  #Last date of midsem break is wrong
    else:
      return 2  #Staring date of midsem exam is wrong
  else:
    return 3  #Ending date of semester is wrong


def find_day_of_week(date_string):
  """
    Finds the day of the week for a given date string.

    Parameters:
    - date_string (str): A string representing a date in the format "%Y%m%d".

    Returns:
    - day_name (str): The name of the day of the week for the given date.
      Returns "Invalid date format" if the input date string is not in the correct format.

    The function parses the input date string, calculates the day of the week as an integer
    (where 0 represents Monday and 6 represents Sunday), and returns the corresponding day name.

    Usage:
    day_name = find_day_of_week('20231117')
    if day_name != "Invalid date format":
        # Process the day_name variable
    else:
        # Handle the case where the date format is invalid

    """
  try:
    # Parse the input date string
    date = datetime.strptime(date_string, "%Y%m%d")

    # Get the day of the week as an integer (0 = Monday, 6 = Sunday)
    day_of_week = date.weekday()

    # Define a list of weekday names
    weekdays = [
        "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
        "Sunday"
    ]

    # Get the weekday name for the given date
    day_name = weekdays[day_of_week]

    return day_name
  except ValueError:
    return "Invalid date format"


# def read_calender():
#   """
#     Reads calendar information from a CSV file based on specific criteria.

#     Returns:
#     - read_calendar_error (int): An error indicator (0 if successful, 1 if there is an error).

#     The function reads the content of a CSV file, specifically looking for two lines:
#     1. "TT: Adjusted Days": Indicates a flag for adjusted days.
#     2. "H:This includes Saturdays/Sundays and GH": Indicates a flag for special inclusions.

#     If both flags are found, the function extracts and returns the relevant information.

#     If the flags are not found or not in the correct form, an error message is printed,
#     and read_calendar_error is set to 1.

#     Usage:
#     read_calendar_error = read_calendar('your_calendar_file.csv')
#     if read_calendar_error == 0:
#         # Process the calendar information
#     else:
#         # Handle the error

#     """
#   read_calender_error = 0
#   # Open the CSV file for the first pass to identify flags
#   with open(filename, 'r') as csvfile:
#     csvreader = csv.reader(csvfile)
#     tt_add_flag = 0
#     this_include_flag = 0
#     # Iterate through each line in the CSV file
#     for i in csvreader:
#       for j in i:
#         if (j.strip().lower() == "tt: adjusted days"):
#           tt_add_flag = 1

#         if (j.strip().lower() == "h:this includes saturdays/sundays and gh"):
#           this_include_flag = 1

#   with open(filename, 'r') as csvfile:
#     line_number = 1
#     reach = 0
#     csvreader = csv.reader(csvfile)
#     find = 0
#     tt_not_take = 0

#     # If both flags are found, extract relevant information
#     if (tt_add_flag == 1 and this_include_flag == 1):
#       for i in csvreader:

#         for j in i:
#           if (j.strip().lower() == "tt: adjusted days" and reach == 0):
#             reach = 1
#             continue
#           if (j.strip().lower() == "h:this includes saturdays/sundays and gh"):
#             find = 1
#             break
#         if (find == 1):
#           break
#         if (reach == 1):

#           # Append relevant information to the 'all' list (assuming 'all' is defined globally)
#           all.append(i)

#     else:
#       print("\nERROR FROM USER INPUT")
#       print(
#           "\nNot written TT: Adjusted Days or H:This includes Saturdays/Sundays and GH in correct form(spelling mistake)"
#       )
#       print("\n")
#       read_calender_error = 1

#   return read_calender_error

# def work_all_calender_data():
#   """
#   Process calendar data stored in the 'all' list and update the 'substitute' and 'holiday' dictionaries.

#   Returns:
#   - error_flag (int): An error indicator (0 if successful, 1 if there is an error).

#   The function iterates through the 'all' list, extracting information about adjusted days and holidays.
#   It updates the 'substitute' dictionary with information about adjusted days
#   and the 'holiday' dictionary with information about holidays.

#   Usage:
#   error_flag = work_all_calendar_data()
#   if error_flag == 0:
#       # Process the 'substitute' and 'holiday' dictionaries
#   else:
#        # Handle the error

#   """
#   error_flag = 0
#   l = 0
#   for i in all:
#     if (l == 0):
#       l = 1
#       continue
#     if (i[0] != ""):
#       a = i[0]
#       if "-" not in a:
#         error_flag = 1
#         #print("came")
#         print(
#             "\nError in the academic calendar ->  in the substitue date format"
#         )
#         print("\nyou have missed - in the input -> ", a)
#         print("\nplease write it correct\n")
#         continue
#       s = a.split("-")
#       day = s[0].split()
#       on_which = s[1].split()
#       # print(s)
#       # print(day[1])
#       # print(day[0])
#       temp1 = day[0]
#       match = re.match(r'(\d+)([a-zA-Z]+)', temp1)
#       number = match.group(1)
#       characters = match.group(2)
#       final_Number = ""
#       if (len(number) == 1):
#         final_Number = "0" + number
#       else:
#         final_Number = number
#         #print(final_Number)
#       final_month = ""
#       lower_month = day[1].strip()
#       if lower_month.lower() not in check_month:

#         print(
#             "\nError in the academic calendar ->  in the substitue date format"
#         )
#         print(
#             "\nYour have written wrong spelling of the month as per given the document in the substitute please check in -> ",
#             a)
#         print("\nplease write it correct\n")
#         error_flag = 1
#         # continue

#       if lower_month.lower() in month_captial_written.keys():
#         final_month = month_captial_written[lower_month.lower()]

#       elif lower_month.lower() in month_small_written.keys():
#         final_month = month_small_written[lower_month.lower()]

#       create_day = str(year) + final_month + str(final_Number)
#       nn = len(create_day)
#       #create_day = create_day[:nn-2]
#       d = str()

#       lower_day = on_which[0].strip()
#       if lower_day.lower() not in check_days:
#         print(
#             "\nError in the academic calendar ->  in the substitue date format"
#         )
#         print(
#             "\nYour have written wrong spelling of the day as per given the document in the substitute please check in ->",
#             a)
#         print("\nplease write it correct\n")
#         error_flag = 1
#         # continue

#       on_what = ""

#       if lower_day.lower() in days.keys():
#         on_what = days[lower_day.lower()]

#       else:
#         on_what = lower_day.lower()

#         # print("Look-1")
#         # print(create_day)
#         # print(on_what)
#         # print(subsitute)
#       subsitute[create_day] = on_what
#       on_what1 = find_day_of_week(create_day)
#       holiday[create_day] = on_what1.lower()
#       #print("Look-1#")

#     length = len(i) - 1
#     start = 1

#     while (start <= length):
#       if (i[start] == ""):
#         start += 1
#         continue
#       elif (i[start].strip().lower() ==
#             "*mid recess & summer vacation - for ug students only"):
#         start += 1
#         continue
#       else:
#         # print("Here-2")
#         if "-" not in i[start] or "," not in i[start + 1]:
#           print(
#               "\nError in the academic calendar ->  in the holiday date format"
#           )
#           print("\nYou have missed - or , in the input ->",
#                 i[start] + " " + i[start + 1])
#           print("\nplease correct it\n")
#           error_flag = 1
#           start += 2
#           continue

#         x = i[start].split("-")
#         y = i[start + 1].split(",")
#         lower_month = x[1].strip()
#         if lower_month.lower() not in check_month:
#           print(
#               "\nError in the academic calendar ->  in the holiday date format"
#           )
#           print(
#               "\nYour have written wrong spelling of the month as per given the document in the substitute please check ->",
#               i[start] + " " + i[start + 1])
#           print("\nplease correct it\n")
#           error_flag = 1
#           start += 2
#           # continue
#           # print(start)
#           # print(x)
#         temp1 = str(x[0])
#         final_Number = ""
#         if (len(temp1) == 1):
#           final_Number = "0" + temp1
#         else:
#           final_Number = temp1

#         hol_day_number = ""
#         if lower_month.lower() in month_captial_written.keys():
#           hol_day_number = month_captial_written[lower_month.lower()]
#         elif lower_month.lower() in month_small_written.keys():
#           hol_day_number = month_small_written[lower_month.lower()]

#           # print(final_Number)
#           # print(y)

#         holiday_date = str(year) + hol_day_number + final_Number
#         lower_day = y[1].strip()
#         if lower_day.lower() not in check_days:
#           print(
#               "\nError in the academic calendar ->  in the holiday date format"
#           )
#           print(
#               "\nYour have written wrong spelling of the day as per given the document in the substitute please check ->",
#               i[start] + " " + i[start + 1])
#           print("\nplease correct it\n")
#           error_flag = 1
#           start += 2
#           # continue
#           # print(mm[x[1]])
#           # print(holiday_date)
#         holiday_date = holiday_date
#         holiday_day = ""
#         if lower_day.lower() in days.keys():
#           holiday_day = days[lower_day.lower()]

#         else:
#           holiday_day = lower_day.lower()

#           #print(holiday_date)
#         holiday[holiday_date] = holiday_day
#         start += 2


#   return error_flag
def is_string_in_list(input_string, string_list):
  return input_string in string_list


def check_days_format(day_dict):
  check_days = [
      "mon", "monday", "tues", "tuesday", "wed", "wednesday", "thur",
      "thursday", "fri", "friday", "sunday", "sun", "sat", "saturday"
  ]

  for datee, day in day_dict.items():
    # Convert the day to lowercase for case-insensitive comparison
    day_lower = day.lower()

    # Check if the day is in the correct format
    if (is_string_in_list(day_lower, check_days)):
      continue
    else:
      print(f"Invalid day format for date {datee}: {day_lower}")
      return False

  # If all days are in the correct format
  return True


def read_work_calender():
  error_flag = 0
  
  csv_file_path = input("Enter your list of all holidays in csv format: ")
  with open(csv_file_path, 'r') as csvfile:
    csvreader = csv.reader(csvfile)
    for i in csvreader:
      if (i[0].strip().lower() == 'date'):
        continue
      date_components = i[0].split('/')
      # print(date_components)
      # print("----------")
      if (len(date_components[0]) == 1):
        date_components[0] = "0" + date_components[0]
        # print(date_components[0])
        # print("&&&&&&&&&&&&")
      if (len(date_components[1]) == 1):
        date_components[1] = "0" + date_components[1]
        #print(date_components)
      formatted_date = ''.join(
          [date_components[2], date_components[1], date_components[0]])
      holiday.update({formatted_date: i[1].lower()})

  # Read the CSV file
  csv_file_path = input("Enter your list of subsituted days in csv format: ")

  with open(csv_file_path, 'r') as csvfile:
    csvreader = csv.reader(csvfile)
    for i in csvreader:
      if (i[0].strip().lower() == 'date'):
        continue
      date_components = i[0].split('/')
      # print(date_components)
      # print("----------")
      if (len(date_components[0]) == 1):
        date_components[0] = "0" + date_components[0]
        # print(date_components[0])
        # print("&&&&&&&&&&&&")
      if (len(date_components[1]) == 1):
        date_components[1] = "0" + date_components[1]
        #print(date_components)
      formatted_date = ''.join(
          [date_components[2], date_components[1], date_components[0]])
      subsitute.update({formatted_date: i[1].lower()})
      i[1] = find_day_of_week(formatted_date)
      holiday.update({formatted_date: i[1].lower()})

  if(check_days_format(holiday)):
    error_flag = 0
  else:
    error_flag = 1

  
  if(check_days_format(subsitute)):
    error_flag = 0
  else:
    error_flag = 1
    
  return error_flag

def takingInput(t):
  """
    Displays a date selector window using Tkinter and returns the selected date.

    Parameters:
    - t (str): The message to be displayed above the date selector.

    Returns:
    - selected_date (str): The selected date in the format "%d/%m/%Y".

    The function creates a Tkinter window with a message, a calendar widget, and a submit button.
    After the user selects a date and clicks the submit button, the window is closed, and the selected date is returned.

    Usage:
    selected_date = taking_input('Select a Date')
    print(f"Selected Date: {selected_date}")

    """

  def get_selected_date():
    selected_date = cal.get_date()
    formatted_date = selected_date.strftime("%d/%m/%Y")
    result_var.set(formatted_date)
    window.destroy()  # Close the input window

  # Create the main window
  window = tk.Tk()
  window.title("Date Selector")

  # Create and place the label
  label_message = tk.Label(window, text=t)
  label_message.grid(row=0, column=0, padx=10, pady=10)

  # Create and place the calendar widget
  cal = DateEntry(window,
                  width=12,
                  background="darkblue",
                  foreground="white",
                  borderwidth=2,
                  year=2023,
                  month=11,
                  day=23)
  cal.grid(row=1, column=0, padx=10, pady=10)

  # Create a StringVar to store the result
  result_var = tk.StringVar()

  # Create and place the button
  button_submit = tk.Button(window, text="Submit", command=get_selected_date)
  button_submit.grid(row=2, column=0, pady=10)

  # Run the main loop
  window.mainloop()

  # Return the result after the window is destroyed
  return result_var.get()


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']

#Check-1
# slot = {"4":[["9:30:00","11:00:00",90,"MONDAY"],["9:30:00","11:00:00",90,"WEDNESDAY"],[]],

#         "1":[["9:30:00","11:00:00",90,"TUESDAY"],["9:30:00","11:00:00",90,"THURSDAY"]],

#         "10":[["18:00:00","19:30:00",90,"MONDAY"],["18:00:00","19:30:00",90,"WEDNESDAY"]],

#         "2":[["11:00:00","12:30:00",90,"MONDAY"],["11:00:00","12:30:00",90,"THURSDAY"]],

#         "5":[["11:00:00","12:30:00",90,"TUESDAY"],["11:00:00","12:30:00",90,"FRIDAY"]],

#         "9":[["11:00:00","12:30:00",90,"WEDNESDAY"],["9:30:00","11:00:00",90,"FRIDAY"]],

#         "6":[["15:00:00","16:30:00",90,"MONDAY"],["15:00:00","16:30:00",90,"WEDNESDAY"]],

#         "3":[["15:00:00","16:30:00",90,"TUESDAY"],["15:00:00","16:30:00",90,"THURSDAY"]],

#         "7":[["16:30:00","18:00:00",90,"MONDAY"],["16:30:00","18:00:00",90,"WEDNESDAY"]],

#         "8":[["16:30:00","18:00:00",90,"TUESDAY"],["16:30:00","18:00:00",90,"THURSDAY"]],
#         }


def read_slotdata_file():
  """
    Reads slot data from a CSV file and populates the 'slot' dictionary.

    Returns:
    - read_slot_data_error (int): An error indicator (0 if successful, 1 if there is an error).

    The function prompts the user for the slot data file name, reads the file,
    and populates the 'slot' dictionary based on the specified structure.

    Usage:
    read_slot_data_error = read_slotdata_file()
    if read_slot_data_error == 0:
        # Process the 'slot' dictionary
    else:
        # Handle the error

    """
  read_slot_data_error = 0
  SlotFileName = input(
      "Please enter the file location of slotData.csv(which contains list of all slots): "
  )
  with open(SlotFileName, 'r') as read_slot:
    slot_reader = csv.reader(read_slot)
    flag = 0
    for i in slot_reader:

      if i[0].strip().lower() != 'slot no.' and flag == 0:
        print("\nError in the slot Data")
        print("\nYou have not written slot no. in the row1 column 1\n")
        flag = 1
        read_slot_data_error = 1
        break
      if i[0].strip().lower() == 'slot no.':
        flag = 1
        continue
      if (len(i[0]) == 0):
        continue
      else:
        temp_list = []  # Create a new list for each row of data
        if (len(i[0]) == 0 or len(i[1]) == 0 or len(i[2]) == 0
            or len(i[3]) == 0):
          print("\nError in the slot Data")
          print("\nEmpty column in the slot data here is the data -> ", i)
          print("\n")
          read_slot_data_error = 1
          continue
        if ":" not in i[1] and "." not in i[1]:
          a = i[1]
          i[1] = a + ":00:00"
        if ":" not in i[2] and "." not in i[2]:
          a = i[2]
          i[2] = a + ":00:00"

        temp_list.append(i[1])
        temp_list.append(i[2])
        temp_list.append(90)
        temp_list.append(i[3].upper())

        if i[3].upper() not in capital_days:
          print("\nError in the slot Data")
          print(
              "\nDay spelling error  in the slotdata file here is the error -> ",
              i)
          print("\n")
          read_slot_data_error = 1
          continue

        main_list.append(
            temp_list.copy())  # Use copy() to avoid modifying the same list
        temp_list.clear()

        if (len(i[4]) == 0 or len(i[5]) == 0 or len(i[6]) == 0):
          print("\nError in the slot Data")
          print("\nEmpty column in the slot data here is the data ->", i)
          print("\n")
          read_slot_data_error = 1
          continue
        if ":" not in i[4] and "." not in i[4]:
          a = i[4]
          i[4] = a + ":00:00"
        if ":" not in i[5] and "." not in i[5]:
          a = i[5]
          i[5] = a + ":00:00"

        temp_list.append(i[4])
        temp_list.append(i[5])

        temp_list.append(90)

        if i[6].upper() not in capital_days:
          print("\nError in the slot Data")
          print(
              "\nDay spelling error  in the slotdata file here is the error->",
              i)
          print("\n")
          read_slot_data_error = 1
          continue

        temp_list.append(i[6].upper())
        main_list.append(
            temp_list.copy())  # Use copy() to avoid modifying the same list

        slot[i[0]] = main_list.copy(
        )  # Use copy() to avoid modifying the same list
        temp_list.clear()
        main_list.clear()
  return read_slot_data_error


def reading_rest_file(rest_file):
  """
    Reads rest data from a CSV file and categorizes it into 'not_in_slot' and 'slot_data' lists.

    Parameters:
    - rest_file (str): The name of the CSV file containing rest data.

    Returns:
    - rest_file_error (int): An error indicator (0 if successful, 1 if there is an error).

    The function reads the rest data from the specified CSV file and categorizes it into two lists:
    1. 'not_in_slot': Contains rows with "Nil" in the eighth column.
    2. 'slot_data': Contains rows with valid slot data.

    Usage:
    rest_file_error = reading_rest_file('your_rest_data_file.csv')
    if rest_file_error == 0:
        # Process the 'not_in_slot' and 'slot_data' lists
    else:
        # Handle the error

    """

  filename = rest_file
  rest_file_error = 0
  with open(filename, encoding='windows-1252') as read_rest_file:
    rest_reader = csv.reader(read_rest_file)
    for i in rest_reader:
      a = i[2].strip()
      if (a.lower() != "rest" and a.lower() != "ii year"
          and a.lower() != "i year"):
        continue
      else:
        b = i[8].strip()
        if (b.lower() == "nil"):
          if (len(i[11]) == 0 or len(i[12]) == 0 or len(i[13]) == 0):
            print("\nError in the List_of_courses")
            print(
                "\nEmpty column(time/day) in the List_of_courses here is the data ->",
                i)
            continue
          if i[13].strip().lower() not in check_days:
            print("\nError in the List_of_courses")
            print(
                "\nSpelling of the day in the List_of_courses is wrong here is the row where is the error->",
                i)
            print("\n")
            rest_file_error = 1
          not_in_slot.append(i)
        elif (len(b) == 0):
          continue
        else:
          slot_data.append(i)

    return rest_file_error


def read_first_year():
  """
    Reads first-year course data from a CSV file and populates the 'first_year_data' dictionary.

    Returns:
    - read_first_year_error_flag (int): An error indicator (0 if successful, 1 if there is an error).

    The function reads first-year course data from the specified CSV file and populates the 'first_year_data' dictionary.
    The dictionary structure is {'course_code': [course_data1, course_data2, ...]}.
    Each 'course_data' entry contains room number, professor name, email, start time, end time, and day.

    Usage:
    read_first_year_error_flag = read_first_year()
    if read_first_year_error_flag == 0:
        # Process the 'first_year_data' dictionary
    else:
        # Handle the error

    """
  read_first_year_error_flag = 0
  filename = input("Enter data of 1st year in csv file:")
  with open(filename, 'r') as read_rest_file:
    rest_reader = csv.reader(read_rest_file)
    flag = 0
    for i in rest_reader:
      if (len(i[0]) == 0):
        continue
      if (flag == 0):
        flag += 1
        continue
      total_class_of_course = int(i[0])
      index_time = 9
      for j in range(total_class_of_course):

        if (i[3] not in first_year_data.keys()):
          course_data = []
          course_data.append(i[6])  #room no
          course_data.append(i[7])  #prof name
          course_data.append(i[8])  #mail
          course_data.append(i[9])  #starting 1 time
          course_data.append(i[10])  #ending 1 time
          course_data.append(i[11])  #day1
          if i[11] not in check_days:
            print(
                "you have written wrong spelling of the day as per given input instruction in the ",
                i)

            read_first_year_error_flag = 1
            continue

          first_year_data[i[3]] = []
          first_year_data[i[3]].append(course_data)
        else:
          for course_data in first_year_data[i[3]]:
            if (course_data[3] == i[index_time]):
              if i[index_time + 2] not in check_days:
                print(
                    "you have written wrong spelling of the day as per given input instruction in the ",
                    i)
                read_first_year_error_flag = 1
                break
              course_data.append(i[index_time + 2])
              break
          else:
            same_course_newtime = []
            same_course_newtime.append(i[6])  #roo no.append(j[7]) #prof name
            same_course_newtime.append(i[7])  #mail
            same_course_newtime.append(i[8])  #mail
            same_course_newtime.append(i[index_time])  #starting 1 time
            same_course_newtime.append(i[index_time + 1])  #ending 1 time
            same_course_newtime.append(i[index_time + 2])  #day1

            if i[index_time + 2] not in check_days:
              print(
                  "you have written wrong spelling of the day as per given input instruction in the ",
                  i)
              read_first_year_error_flag = 1
              continue
            first_year_data[i[3]].append(same_course_newtime)

        index_time += 3
  return read_first_year_error_flag


def get_next_day_occurrence(input_date, target_day):
  """
    Calculates the next occurrence date of a specified day of the week after a given input date.

    Parameters:
    - input_date (str): Input date in the format "%d/%m/%Y".
    - target_day (str): Target day of the week (e.g., "Monday", "Tuesday").

    Returns:
    - next_occurrence_date (str): The next occurrence date in the format "%d/%m/%Y".

    The function takes an input date and a target day of the week, calculates the difference
    between the target day and the current day, and determines the next occurrence date.

    Usage:
    next_occurrence = get_next_day_occurrence('01/01/2023', 'Wednesday')
    print(f"The next Wednesday after 01/01/2023 is on: {next_occurrence}")

    """
  # Convert the input date string to a datetime object
  date_obj = datetime.strptime(input_date, "%d/%m/%Y")

  # Define a mapping for days of the week
  days_mapping = {
      "Monday": 0,
      "Tuesday": 1,
      "Wednesday": 2,
      "Thursday": 3,
      "Friday": 4,
      "Saturday": 5,
      "Sunday": 6
  }

  # Get the numerical representation of the target day
  target_day_num = days_mapping.get(target_day.capitalize())
  if target_day_num is None:
    raise ValueError("Invalid day of the week")

  # Calculate the difference between the target day and the current day
  day_difference = (target_day_num - date_obj.weekday() + 7) % 7

  # Calculate the next occurrence date
  next_occurrence_date = date_obj + timedelta(days=day_difference)

  # Format and return the result
  return next_occurrence_date.strftime("%d/%m/%Y")


# testing =[10,19,29,35,39,49,59]
# for the slot data
def do_mail(s_date, e_date):
  """
    Sends emails based on the course slots, not_in_slot, and first-year data within the specified date range.

    Parameters:
    - s_date (str): Start date in the format "%d/%m/%Y".
    - e_date (str): End date in the format "%d/%m/%Y".

    The function iterates through the slotdata, not_in_slot, and first_year_data, and sends emails accordingly.
    It uses the mail, mail_2, mail_not_in_slot, and mail_first_year functions to handle the email sending logic.

    Usage:
    do_mail('dd/mm/yyyy', 'dd/mm/yyyy')

    """
  #print("Here-1")
  for allCourses in range(len(slot_data)):  #len(slot_data)
    # allCourses=105
    print("\n")
    print(allCourses)
    print(slot_data[allCourses])
    course_list = slot_data[
        allCourses]  #['CSE', 'CSE506', 'Rest', 'Data Mining', 'DMG', '40', 'C215', 'Vikram Goyal', '2', '2', '', '', '', '',sample@iiitd.ac.in]
    #print(course_list)
    course_slot = course_list[8]  #2
    if (course_slot == "" or course_slot == " "):
      continue
    if course_slot not in slot.keys():
      continue
    slot_list = slot[
        course_slot]  #[['11:00:00', '12:30:00', 90, 'MONDAY'], ['11:00:00', '12:30:00', 90, 'THURSDAY']]
    # print(course_slot)
    # print(slot_list)
    # time starting time slot_d1 and slot_d2
    slot_d1 = slot_list[0][0]
    slot_d2 = slot_list[1][0]
    # print(slot_d1)
    # print(slot_d2)

    if (slot_d1 == slot_d2):
      mail(s_date, e_date, allCourses)
    
    else:
      mail_2(s_date, e_date, allCourses)
  #     #print("Here-1#")c

  for allCourses in range(len(not_in_slot)):
    print("\n")
    print(allCourses)
    print(not_in_slot[allCourses])
    mail_not_in_slot(s_date, e_date, allCourses)

  # mail_firt_year(s_date,e_date)


def mail(s_date, e_date, allCourses):
  """
    Sends a recurring event invitation for a specific course during the specified date range.

    Parameters:
    - s_date (datetime.date): Start date of the recurring event.
    - e_date (datetime.date): End date of the recurring event.
    - allCourses (int): Index of the course in the slot_data list.

    The function uses the Google Calendar API to create a recurring event for the specified course during the given date range.
    It extracts necessary information from the slot_data and slot dictionaries, such as course details, location, time slots,
    and recurrence rules.

    Usage:
    For those courses which have same time on both the days
    mail(datetime.date(dd/mm/yyyy), datetime.date(dd/mm/yyyy))
    """

  #print("In mail")
  course_list = slot_data[
      allCourses]  #['CSE', 'CSE506', 'Rest', 'Data Mining', 'DMG', '40', 'C215', 'Vikram Goyal', '2', '2', '', '', '', '']
  # print(course_list)
  location = course_list[6]
  course = course_list[3]  #Data Mining
  course_slot = course_list[8]  #2
  emali_id = course_list[14]
  #print(course_slot)
  slot_list = slot[
      course_slot]  #[['11:00:00', '12:30:00', 90, 'MONDAY'], ['11:00:00', '12:30:00', 90, 'THURSDAY']]
  #slot_d1/d2 day
  slot_d1 = slot_list[0][3]  #Monday
  slot_d2 = slot_list[1][3]  #Thruesday
  # t1/t2 time
  t1 = slot_list[0][0]
  t2 = slot_list[1][0]
  tt1 = []  #Stores starting time
  tt2 = []  #Stores Ending Time
  if ":" in t1:
    tt1 = t1.split(":")
  if "." in t1:
    tt1 = t1.split(".")

  if ":" in t2:
    tt2 = t2.split(":")
  if "." in t2:
    tt2 = t2.split(".")

  if (len(tt1[0]) == 1):
    tt1[0] = "0" + tt1[0]
  if (len(tt1) == 1):
    tt1.append("00")
  if (len(tt1) == 2):
    tt1.append("00")
  if (len(tt2) == 1):
    tt2.append("00")
  if (len(tt2) == 2):
    tt2.append("00")
  time_d1 = [tt1[0], tt1[1], tt1[2]]
  x = ["05", "30", "00"]
  x_date = timedelta(hours=int(x[0]), minutes=int(x[1]), seconds=int(x[2]))

  timee_d1 = timedelta(hours=int(time_d1[0]),
                       minutes=int(time_d1[1]),
                       seconds=int(time_d1[2]))

  remaing_time = timee_d1 - x_date  #Stores time difference between starting time and 05:30:00(GMT-INDIA)
  remaing_time = str(remaing_time)
  remaing_time = remaing_time.split(":")
  # actual holiday
  ee_date = []
  hol_size = len(hol)
  # print(remaing_time)
  # print(hol)
  # add 0 if need
  for i, j in hol.items():
    # print(i) #20231120
    # print(j) #Friday
    j = j.lower()
    b = j.strip()
    if (b == slot_d1.lower() or b == slot_d2.lower()):
      ee_date.append(i)
  # print("First Here:")
  # print(ee_date)

  rr_date = []  #Actual subsitute day
  sub_size = len(sub)

  for i, j in sub.items():
    j = j.lower()
    b = j.strip()
    # print("DBZ")
    # print(j)
    # print(b)
    # print(slot_d1)
    # print(slot_d2.lower())
    if (b == slot_d1.lower() or b == slot_d2.lower()):
      rr_date.append(i)

  # print(rr_date)
  start_t = slot_list[0][0]
  start_t = tt1
  end_t = slot_list[0][1]
  if ":" in end_t:
    end_t = end_t.split(":")
  if "." in end_t:
    end_t = end_t.split(".")

  if (len(end_t[0]) == 1):
    end_t[0] = "0" + end_t[0]
  if (len(end_t) == 1):
    end_t.append("00")
  if (len(end_t) == 2):
    end_t.append("00")

  #-----(2)
  # if(int(start_t[0]) > 0 and int(start_t[0]) <7):
  #     a = int(start_t[0])+12
  #     start_t[0] = str(a)
  # if(int(end_t[0]) > 0 and int(end_t[0]) <7):
  #     a = int(end_t[0])+12
  #     end_t[0] = str(a)
  #------
  startt_time = start_t[0] + ":" + start_t[1] + ":" + start_t[2]
  endd_time = end_t[0] + ":" + end_t[1] + ":" + end_t[2]

  datee = s_date.strftime("%d")
  mont = s_date.strftime("%m")
  yea = s_date.strftime("%Y")

  s_date = datee + "/" + mont + "/" + yea
  next_occurrence = get_next_day_occurrence(s_date, slot_list[0][3])
  n_o = next_occurrence.split("/")
  s_date = date(int(n_o[2]), int(n_o[1]), int(n_o[0]))

  # s_date = s_date + timedelta(days=day_number[slot_list[0][3]])
  datee = s_date.strftime("%d")
  mont = s_date.strftime("%m")
  yea = s_date.strftime("%Y")

  start_datetime = yea + "-" + mont + "-" + datee + "T" + startt_time + "+05:30"
  end_datetime = yea + "-" + mont + "-" + datee + "T" + endd_time + "+05:30"

  datee1 = e_date.strftime("%d")
  mont1 = e_date.strftime("%m")
  yea1 = e_date.strftime("%Y")
  s_d1 = short_term[slot_d1]
  s_d2 = short_term[slot_d2]
  aa = s_d1 + "," + s_d2
  rrule = 'RRULE:FREQ=WEEKLY;UNTIL=' + yea1 + mont1 + datee1 + ';BYDAY=' + aa  #RR rule

  add_remaing = remaing_time[0] + remaing_time[1] + remaing_time[2]
  # print("Hello-1")
  # print(add_remaing)
  if (len(add_remaing) == 5):
    add_remaing = "0" + add_remaing + "Z"
    # 040000Z
  else:
    add_remaing = add_remaing + "Z"

  exdate = ""  #EXDATE Rule
  if (len(ee_date) != 0):
    exdate = "EXDATE:"
  len_eedate = len(ee_date)
  ac = 0
  for i in ee_date:

    len_eedate -= 1

    exdate = exdate + i + "T" + add_remaing

    if (len_eedate != 0):
      exdate += ","
    #print(exdate)

  rdate = ""  #Rdate Rule
  if (len(rr_date) != 0):
    rdate = "RDATE:"
  # print("JJK")
  # print("email= "+ emali_id)
  # print(rdate)
  len_rdate = len(rr_date)
  for i in rr_date:
    len_rdate -= 1
    rdate = rdate + i + "T" + add_remaing
    if (len_rdate != 0):
      rdate += ","
  # print("Here-1")
  # print(rdate)
  # print(start_datetime)
  # print(end_datetime)
  # print(rrule)
  #print(exdate)
  """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file('credentials.json',
                                                       SCOPES)
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
      token.write(creds.to_json())

  try:
    service = build('calendar', 'v3', credentials=creds)
    event = {
        'summary': course,
        'location': location,
        'description': 'Happy Learning...:)',
        'start': {
            'dateTime': start_datetime,
            'timeZone': 'Asia/Kolkata'
        },
        'end': {
            'dateTime': end_datetime,
            'timeZone': 'Asia/Kolkata'
        },
        'recurrence': [
            rrule,
            rdate,
            exdate,
        ],
        'attendees': [{
            'email': emali_id
        }],
    }

    # Insert the event
    event = service.events().insert(calendarId='primary', body=event).execute()
    print(f'Recurring event created: {event["htmlLink"]}')
  except HttpError as error:
    print('An error occurred: %s' % error)


def mail_2(s_date, e_date, allCourses):
  """
    Sends a recurring event invitation for a specific course during the specified date range.

    Parameters:
    - s_date (datetime.date): Start date of the recurring event.
    - e_date (datetime.date): End date of the recurring event.
    - allCourses (int): Index of the course in the slot_data list.

    The function uses the Google Calendar API to create a recurring event for the specified course during the given date range.
    It extracts necessary information from the slot_data and slot dictionaries, such as course details, location, time slots,
    and recurrence rules.

    Usage:
    For those courses which have different timings of their classes on both days
    mail(datetime.date(2023, 1, 1), datetime.date(2023, 1, 15), 0)
    """

  #General information of the course(same as above function i.e mail)
  course_list = slot_data[allCourses]
  course = course_list[3]
  location = course_list[6]
  emali_id = course_list[14]
  course_slot = course_list[8]
  slot_list = slot[course_slot]
  slot_d1 = slot_list[0][3]
  slot_d2 = slot_list[1][3]
  t1 = slot_list[0][0]
  t2 = slot_list[1][0]
  tt1 = []
  tt2 = []
  if ":" in t1:
    tt1 = t1.split(":")
  if "." in t1:
    tt1 = t1.split(".")

  if ":" in t2:
    tt2 = t2.split(":")
  if "." in t2:
    tt2 = t2.split(".")

  if (len(tt1) == 1):
    tt1.append("00")
  if (len(tt2) == 1):
    tt2.append("00")
  if (len(tt1) == 2):
    tt1.append("00")
  if (len(tt2) == 2):
    tt2.append("00")

  if (len(tt1[0]) == 1):
    tt1[0] = "0" + tt1[0]
  if (len(tt2[0]) == 1):
    tt2[0] = "0" + tt2[0]

  time_d1 = [tt1[0], tt1[1], tt1[2]]
  time_d2 = [tt2[0], tt2[1], tt2[2]]
  x = ["05", "30", "00"]
  x_date = timedelta(hours=int(x[0]), minutes=int(x[1]), seconds=int(x[2]))

  timee_d1 = timedelta(hours=int(time_d1[0]),
                       minutes=int(time_d1[1]),
                       seconds=int(time_d1[2]))
  timee_d2 = timedelta(hours=int(time_d2[0]),
                       minutes=int(time_d2[1]),
                       seconds=int(time_d2[2]))

  remaing_time1 = timee_d1 - x_date
  remaing_time1 = str(remaing_time1)
  remaing_time1 = remaing_time1.split(":")

  remaing_time2 = timee_d2 - x_date
  remaing_time2 = str(remaing_time2)
  remaing_time2 = remaing_time2.split(":")
  #print(remaing_time1)
  ee_date_1 = []
  ee_date_2 = []
  hol_size = len(hol)
  for i, j in hol.items():
    j = j.lower()
    b = j.strip()
    if (b == slot_d1.lower()):
      ee_date_1.append(i)
    elif (b == slot_d2.lower()):
      ee_date_2.append(i)

  rr_date_1 = []
  rr_date_2 = []
  sub_size = len(sub)
  for i, j in sub.items():
    j = j.lower()
    b = j.strip()
    if (b == slot_d1.lower()):
      rr_date_1.append(i)

    elif (b == slot_d2.lower()):
      rr_date_2.append(i)

  #Mail for the first day of course

  datee = s_date.strftime("%d")
  mont = s_date.strftime("%m")
  yea = s_date.strftime("%Y")

  s_date_1 = datee + "/" + mont + "/" + yea
  next_occurrence = get_next_day_occurrence(s_date_1, slot_list[0][3])
  n_o = next_occurrence.split("/")
  s_date_1 = date(int(n_o[2]), int(n_o[1]), int(n_o[0]))

  # s_date_1 = s_date + timedelta(days=day_number[slot_list[0][3]])
  datee = s_date_1.strftime("%d")
  mont = s_date_1.strftime("%m")
  yea = s_date_1.strftime("%Y")

  start_t1 = slot_list[0][0]
  start_t1 = tt1
  end_t1 = slot_list[0][1]
  if ":" in end_t1:
    end_t1 = end_t1.split(":")
  if "." in end_t1:
    end_t1 = end_t1.split(".")

  if (len(end_t1[0]) == 1):
    end_t1[0] = "0" + end_t1[0]
  if (len(end_t1) == 1):
    end_t1.append("00")
  if (len(end_t1) == 2):
    end_t1.append("00")

  # if(int(start_t1[0]) > 0 and int(start_t1[0]) <7):
  #     a = int(start_t1[0])+12
  #     start_t1[0] = str(a)
  # if(int(end_t1[0]) > 0 and int(end_t1[0]) <7):
  #     a = int(end_t1[0])+12
  #     end_t1[0] = str(a)
  startt_time1 = start_t1[0] + ":" + start_t1[1] + ":" + start_t1[2]
  endd_time1 = end_t1[0] + ":" + end_t1[1] + ":" + end_t1[2]
  start_datetime_1 = yea + "-" + mont + "-" + datee + "T" + startt_time1 + "+05:30"
  end_datetime_1 = yea + "-" + mont + "-" + datee + "T" + endd_time1 + "+05:30"
  datee_1 = e_date.strftime("%d")
  mont_1 = e_date.strftime("%m")
  yea_1 = e_date.strftime("%Y")
  s_d1 = short_term[slot_d1]

  rrule1 = 'RRULE:FREQ=WEEKLY;UNTIL=' + yea_1 + mont_1 + datee_1 + ';BYDAY=' + s_d1

  add_remaing1 = remaing_time1[0] + remaing_time1[1] + remaing_time1[2]

  if (len(add_remaing1) == 5):
    add_remaing1 = "0" + add_remaing1 + "Z"
  else:
    add_remaing1 = add_remaing1 + "Z"
  exdate1 = ""
  if (len(ee_date_1) != 0):
    exdate1 = "EXDATE:"
  len_eedate1 = len(ee_date_1)
  ac = 0
  for i in ee_date_1:
    len_eedate1 -= 1
    exdate1 = exdate1 + i + "T" + add_remaing1
    if (len_eedate1 != 0):
      exdate1 += ","

  rdate1 = ""
  if (len(rr_date_1) != 0):
    rdate1 = "RDATE:"
  len_rdate1 = len(rr_date_1)
  for i in rr_date_1:
    len_rdate1 -= 1
    rdate1 = rdate1 + i + "T" + add_remaing1
    if (len_rdate1 != 0):
      rdate1 += ","




  
  """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file('credentials.json',
                                                       SCOPES)
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
      token.write(creds.to_json())

  try:
    service = build('calendar', 'v3', credentials=creds)
    event = {
        'summary': course,
        'location': location,
        'description': 'Happy Learning...:)',
        'start': {
            'dateTime': start_datetime_1,
            'timeZone': 'Asia/Kolkata'
        },
        'end': {
            'dateTime': end_datetime_1,
            'timeZone': 'Asia/Kolkata'
        },
        'recurrence': [
            rrule1,
            exdate1,
            rdate1,
        ],
        'attendees': [{
            'email': emali_id
        }],
    }

    # Insert the event
    event = service.events().insert(calendarId='primary', body=event).execute()
    print(f'Recurring event created: {event["htmlLink"]}')
  except HttpError as error:
    print('An error occurred: %s' % error)

  # mail for 2nd day of course
  datee = s_date.strftime("%d")
  mont = s_date.strftime("%m")
  yea = s_date.strftime("%Y")

  s_date_2 = datee + "/" + mont + "/" + yea
  next_occurrence = get_next_day_occurrence(s_date_2, slot_list[1][3])
  n_o = next_occurrence.split("/")
  s_date_2 = date(int(n_o[2]), int(n_o[1]), int(n_o[0]))

  # s_date_2 = s_date + timedelta(days=day_number[slot_list[1][3]])
  datee = s_date_2.strftime("%d")
  mont = s_date_2.strftime("%m")
  yea = s_date_2.strftime("%Y")

  start_t2 = slot_list[1][0]
  start_t2 = tt2
  end_t2 = slot_list[1][1]
  print(slot_list)
  if ":" in end_t2:
    end_t2 = end_t2.split(":")
  if "." in end_t2:
    end_t2 = end_t2.split(".")

  if (len(end_t2[0]) == 1):
    end_t2[0] = "0" + end_t2[0]
  if (len(end_t2) == 1):
    end_t2.append("00")
  if (len(end_t2) == 2):
    end_t2.append("00")
  # if(int(start_t2[0]) > 0 and int(start_t2[0]) <7):
  #     a = int(start_t2[0])+12
  #     start_t2[0] = str(a)
  # if(int(end_t2[0]) > 0 and int(end_t2[0]) <7):
  #     a = int(end_t2[0])+12
  #     end_t2[0] = str(a)

  startt_time2 = start_t2[0] + ":" + start_t2[1] + ":" + start_t2[2]
  endd_time2 = end_t2[0] + ":" + end_t2[1] + ":" + end_t2[2]
  start_datetime_2 = yea + "-" + mont + "-" + datee + "T" + startt_time2 + "+05:30"
  end_datetime_2 = yea + "-" + mont + "-" + datee + "T" + endd_time2 + "+05:30"
  datee_2 = e_date.strftime("%d")
  mont_2 = e_date.strftime("%m")
  yea_2 = e_date.strftime("%Y")
  s_d2 = short_term[slot_d2]
  rrule2 = 'RRULE:FREQ=WEEKLY;UNTIL=' + yea_2 + mont_2 + datee_2 + ';BYDAY=' + s_d2
  add_remaing2 = remaing_time2[0] + remaing_time2[1] + remaing_time2[2]
  if (len(add_remaing2) == 5):
    add_remaing2 = "0" + add_remaing2 + "Z"

  exdate2 = ""
  if (len(ee_date_2) != 0):
    exdate2 = "EXDATE:"
  len_eedate2 = len(ee_date_2)
  ac = 0
  for i in ee_date_2:
    len_eedate2 -= 1
    exdate2 = exdate2 + i + "T" + add_remaing2
    if (len_eedate2 != 0):
      exdate2 += ","
  #print(exdate2)
  rdate2 = ""
  if (len(rr_date_2) != 0):
    rdate2 = "RDATE:"
  len_rdate2 = len(rr_date_2)
  for i in rr_date_2:
    len_rdate2 -= 1
    rdate2 = rdate2 + i + "T" + add_remaing2
    if (len_rdate2 != 0):
      rdate2 += ","


  """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file('credentials.json',
                                                       SCOPES)
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
      token.write(creds.to_json())

  try:
    service = build('calendar', 'v3', credentials=creds)
    event = {
        'summary': course,
        'location': location,
        'description': 'Happy Learning...:)',
        'start': {
            'dateTime': start_datetime_2,
            'timeZone': 'Asia/Kolkata'
        },
        'end': {
            'dateTime': end_datetime_2,
            'timeZone': 'Asia/Kolkata'
        },
        'recurrence': [
            rrule2,
            exdate2,
            rdate2,
        ],
        'attendees': [{
            'email': emali_id
        }],
    }

    # Insert the event
    event = service.events().insert(calendarId='primary', body=event).execute()
    print(f'Recurring event created: {event["htmlLink"]}')
  except HttpError as error:
    print('An error occurred: %s' % error)


# for the courses that doesnt fall in any slot(i.e Slot=Nil)


def mail_not_in_slot(s_date, e_date, index):
  """
    Sends recurring event invitations for a course not in a regular time slot during the specified date range.

    Parameters:
    - s_date (datetime.date): Start date of the recurring event.
    - e_date (datetime.date): End date of the recurring event.
    - index (int): Index of the course in the not_in_slot list.

    The function creates a recurring event for a course not in a regular time slot, considering holidays and substitute dates.
    It uses the Google Calendar API to send invitations for the specified time slot.

    Usage:
    For the courses that doesnt fall in any slot(i.e Slot=Nil)
    mail_not_in_slot(datetime.date(yyyy,mm,dd), datetime.date(yyyy,mm, dd))
    """
  course_list = not_in_slot[
      index]  #list of the course which contains all its details
  course = course_list[3]  #Name of the course
  emali_id = course_list[14]  #Gmail of Course
  #print(course_list)
  course_day = course_list[13].lower()  #Day on which Course lies
  s_time = course_list[11]  #Starting Time
  e_time = course_list[12]  #Ending time

  if "." not in e_time and ":" not in e_time:
    e_time = e_time + ":00"
  if "." not in s_time and ":" not in s_time:
    s_time = s_time + ":00"

  if "." in s_time:
    s_time = s_time.split(".")
  if ":" in s_time:
    s_time = s_time.split(":")

  if "." in e_time:
    e_time = e_time.split(".")
  if ":" in e_time:
    e_time = e_time.split(":")

  if (len(s_time[0]) == 1):
    s_time[0] = "0" + s_time[0]
  if (len(e_time[0]) == 1):
    e_time[0] = "0" + e_time[0]

  if (len(s_time) == 1):
    s_time.append("00")
  if (len(s_time) == 2):
    s_time.append("00")

  if (len(e_time) == 1):
    e_time.append("00")
  if (len(e_time) == 2):
    e_time.append("00")

  if (len(s_time[1]) == 1):
    s_time[1] = s_time[1] + "0"
  if (len(e_time[1]) == 1):
    e_time[1] = e_time[1] + "0"

  start_time = [s_time[0], s_time[1],
                s_time[2]]  #Storing starting time in list
  end_time = [e_time[0], e_time[1], e_time[2]]  #Storing ending time in list

  startt_time = start_time[0] + ":" + start_time[1] + ":" + start_time[2]
  endd_time = end_time[0] + ":" + end_time[1] + ":" + end_time[2]
  x = ["05", "30", "00"]

  x_date = timedelta(hours=int(x[0]), minutes=int(x[1]), seconds=int(x[2]))

  time_d = timedelta(hours=int(start_time[0]),
                     minutes=int(start_time[1]),
                     seconds=int(start_time[2]))
  remaing_time = time_d - x_date  #Difference of the time between starting time and 05:30:00(GMT-INDIA)
  remaing_time = str(remaing_time)
  remaing_time = remaing_time.split(":")

  #Actual Holiday
  ee_date = []
  hol_size = len(hol)

  for i, j in hol.items():
    j = j.lower()
    b = j.strip()
    if (b == course_day.lower()):
      ee_date.append(i)

  #Actual subsitute Day
  rr_date = []
  sub_size = len(sub)

  for i, j in sub.items():
    j = j.lower()
    b = b.strip()
    if (b == course_day.lower()):
      rr_date.append(i)
  datee = s_date.strftime("%d")
  mont = s_date.strftime("%m")
  yea = s_date.strftime("%Y")

  ss_date = datee + "/" + mont + "/" + yea
  next_occurrence = get_next_day_occurrence(ss_date, course_day)
  n_o = next_occurrence.split("/")
  s_date = date(int(n_o[2]), int(n_o[1]), int(n_o[0]))

  datee = s_date.strftime("%d")
  mont = s_date.strftime("%m")
  yea = s_date.strftime("%Y")
  start_datetime = yea + "-" + mont + "-" + datee + "T" + startt_time + "+05:30"
  end_datetime = yea + "-" + mont + "-" + datee + "T" + endd_time + "+05:30"

  datee1 = e_date.strftime("%d")
  mont1 = e_date.strftime("%m")
  yea1 = e_date.strftime("%Y")
  rrule = 'RRULE:FREQ=WEEKLY;UNTIL=' + yea1 + mont1 + datee1 + ';BYDAY=' + short_term[
      course_day.upper()]  #RRule

  add_remaing = remaing_time[0] + remaing_time[1] + remaing_time[2]
  if (len(add_remaing) == 5):
    add_remaing = "0" + add_remaing + "Z"
    # 040000Z

  exdate = ""  #EXDATE
  if (len(ee_date) != 0):
    exdate = "EXDATE:"
  len_eedate = len(ee_date)
  ac = 0
  for i in ee_date:

    len_eedate -= 1

    exdate = exdate + i + "T" + add_remaing
    if (len_eedate != 0):
      exdate += ","

  rdate = ""  #RDATE
  if (len(rr_date) != 0):
    rdate = "RDATE:"
  len_rdate = len(rr_date)
  for i in rr_date:
    len_rdate -= 1
    rdate = rdate + i + "T" + add_remaing
    if (len_rdate != 0):
      rdate += ","
  """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file('credentials.json',
                                                       SCOPES)
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
      token.write(creds.to_json())

  try:
    service = build('calendar', 'v3', credentials=creds)
    event = {
        'summary': course,
        'location': 'Somewhere',
        'description': 'somewhere online',
        'start': {
            'dateTime': start_datetime,
            'timeZone': 'Asia/Kolkata'
        },
        'end': {
            'dateTime': end_datetime,
            'timeZone': 'Asia/Kolkata'
        },
        'recurrence': [rrule, exdate, rdate],
        'attendees': [{
            'email': emali_id
        }],
    }

    # Insert the event
    event = service.events().insert(calendarId='primary', body=event).execute()
    print(f'Recurring event created: {event["htmlLink"]}')
  except HttpError as error:
    print('An error occurred: %s' % error)


def mail_firt_year(s_date, e_date):
  """
    Sends recurring event invitations for first-year courses during the specified date range.

    Parameters:
    - s_date (datetime.date): Start date of the recurring event.
    - e_date (datetime.date): End date of the recurring event.

    The function iterates over the list of first-year courses, creates recurring events for each course based on its schedule,
    considering holidays and substitute dates. It uses the Google Calendar API to send invitations for the specified time slots.

    Usage:
    For the courses of 1st Year
    mail_first_year(datetime.date(dd,mm, yyyy), datetime.date(dd,mm,yyyy))
    """
  #iterating over the .list of courses of 1st year
  for i in first_year_data:
    course_name = i  #course name
    for j in first_year_data[i]:
      location = j[0]  #classroom location
      email_id = j[2]  #Gmail of the course
      first_date = j[5]  #first day of the class

      start_time = j[3]  #starting time of the class
      end_time = j[4]  #Ending time of the class
      list_day_working_upper = []
      list_day_working_lower = []

      for ind in range(5, len(j)):

        list_day_working_upper.append(j[ind].upper())
        list_day_working_lower.append(j[ind].lower())
      course_first_day = j[5]
      #Spliting the time on ':' or '.' basis
      if (":" in start_time):
        start_time = start_time.split(":")
        if (len(start_time[0]) == 1):
          start_time[0] = "0" + start_time[0]
        if (len(start_time) == 1):
          start_time.append("00")
        if (len(start_time) == 2):
          start_time.append("00")
      elif ("." in start_time):
        start_time = start_time.split(".")
        if (len(start_time[0]) == 1):
          start_time[0] = "0" + start_time[0]
        if (len(start_time) == 1):
          start_time.append("00")
        if (len(start_time) == 2):
          start_time.append("00")
      else:
        if (len(start_time) == 1):
          start_time = "0" + start_time
        start_time = [start_time, "00", "00"]

      if (":" in end_time):
        end_time = end_time.split(":")
        if (len(end_time[0]) == 1):
          end_time[0] = "0" + end_time[0]

        if (len(end_time) == 1):
          end_time.append("00")
        if (len(end_time) == 2):
          end_time.append("00")

      elif ("." in end_time):
        end_time = end_time.split(".")
        if (len(end_time[0]) == 1):
          end_time[0] = "0" + end_time[0]
        if (len(end_time) == 1):
          end_time.append("00")
        if (len(end_time) == 2):
          end_time.append("00")
      else:
        if (len(end_time) == 1):
          end_time = "0" + end_time
        end_time = [end_time, "00", "00"]

      class_start_time = start_time[0] + ":" + start_time[
          1] + ":" + start_time[2]  #storing starting time in list
      class_end_time = end_time[0] + ":" + end_time[1] + ":" + end_time[
          2]  #storing ending timeing list

      x = ["05", "30", "00"]
      x_date = timedelta(hours=int(x[0]), minutes=int(x[1]), seconds=int(x[2]))

      timee_d1 = timedelta(hours=int(start_time[0]),
                           minutes=int(start_time[1]),
                           seconds=int("00"))

      remaing_time = timee_d1 - x_date
      remaing_time = str(remaing_time)
      remaing_time = remaing_time.split(
          ":"
      )  #Time difference between starting time of the course and 05:30:30(GMT-INDIA)

      datee = s_date.strftime("%d")
      mont = s_date.strftime("%m")
      yea = s_date.strftime("%Y")

      ss_date = datee + "/" + mont + "/" + yea
      next_occurrence = get_next_day_occurrence(ss_date, course_first_day)

      n_o = next_occurrence.split("/")

      s_date = date(int(n_o[2]), int(n_o[1]), int(n_o[0]))

      datee = s_date.strftime("%d")
      mont = s_date.strftime("%m")
      yea = s_date.strftime("%Y")
      start_datetime = yea + "-" + mont + "-" + datee + "T" + class_start_time + "+05:30"
      end_datetime = yea + "-" + mont + "-" + datee + "T" + class_end_time + "+05:30"

      datee1 = e_date.strftime("%d")
      mont1 = e_date.strftime("%m")
      yea1 = e_date.strftime("%Y")

      short_termm = ""
      for len_list_day_working in range(0, len(list_day_working_upper)):

        if (len_list_day_working < len(list_day_working_upper) - 1):
          short_termm = short_termm + short_term[
              list_day_working_upper[len_list_day_working]] + ","
        else:
          short_termm = short_termm + short_term[
              list_day_working_upper[len_list_day_working]]

      rrule = 'RRULE:FREQ=WEEKLY;UNTIL=' + yea1 + mont1 + datee1 + ';BYDAY=' + short_termm  #RRule
      print(rrule)
      add_remaing = remaing_time[0] + remaing_time[1] + remaing_time[2]
      if (len(add_remaing) == 5):
        add_remaing = "0" + add_remaing + "Z"
        # 040000Z

      # actual holiday
      ee_date = []
      hol_size = len(hol)

      for i, j in hol.items():

        j = j.lower()
        b = j.strip()
        if (b in list_day_working_lower):
          ee_date.append(i)

      rr_date = []
      sub_size = len(sub)

      for i, j in sub.items():
        j = j.lower()
        b = j.strip()

        if (b in list_day_working_lower):
          rr_date.append(i)

      exdate = ""  #EXDATE
      if (len(ee_date) != 0):
        exdate = "EXDATE:"
      len_eedate = len(ee_date)
      ac = 0
      for i in ee_date:

        len_eedate -= 1

        exdate = exdate + i + "T" + add_remaing

        if (len_eedate != 0):
          exdate += ","
        #print(exdate)

      rdate = ""  #RDATE
      if (len(rr_date) != 0):
        rdate = "RDATE:"
      # print("JJK")
      # print(rdate)
      len_rdate = len(rr_date)
      for i in rr_date:
        len_rdate -= 1
        rdate = rdate + i + "T" + add_remaing
        if (len_rdate != 0):
          rdate += ","
      # print("Here-1")
      # print(rdate)
      # print(start_datetime)
      # print(end_datetime)
      # print(rrule)
      #print(exdate)
      """Shows basic usage of the Google Calendar API.
            Prints the start and name of the next 10 events on the user's calendar.
            # """
      creds = None
      # The file token.json stores the user's access and refresh tokens, and is
      # created automatically when the authorization flow completes for the first
      # time.
      if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
      # If there are no (valid) credentials available, let the user log in.
      if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
          creds.refresh(Request())
        else:
          flow = InstalledAppFlow.from_client_secrets_file(
              'credentials.json', SCOPES)
          creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
          token.write(creds.to_json())

      try:
        service = build('calendar', 'v3', credentials=creds)
        event = {
            'summary': course_name,
            'location': location,
            'description': 'Happy Learning...:)',
            'start': {
                'dateTime': start_datetime,
                'timeZone': 'Asia/Kolkata'
            },
            'end': {
                'dateTime': end_datetime,
                'timeZone': 'Asia/Kolkata'
            },
            'recurrence': [rrule, exdate, rdate],
            'attendees': [{
                'email': email_id
            }],
        }

        # Insert the event
        event = service.events().insert(calendarId='primary',
                                        body=event).execute()
        print(f'Recurring event created: {event["htmlLink"]}')
      except HttpError as error:
        print('An error occurred: %s' % error)


def holiday_add(ms_date, holys_date, mid_date, holy_date):
  """
    Adds holidays to the 'hol' dictionary in the specified date range.

    Parameters:
    - ms_date (list): List representing the current date [day, month, year].
    - holys_date (date): Start date of the holiday period.
    - mid_date (date): Current date to start adding holidays.
    - holy_date (date): End date of the holiday period.

    The function iterates through the date range from mid_date to holy_date, adds each date as a key to the 'hol' dictionary,
    and assigns the corresponding day of the week as the value.

    Usage:
    holiday_add([dd, mm, yyyy], date(dd, mm, yyyy), date(dd, mm, yyyy), date(dd, mm, yyyy))
    """
  # print("CheckPoint-1")
  # print(hol)
  while (mid_date <= holy_date):
    # print(mid_date)     #2023-10-28
    # print(ms_date)      #[28,10,2023]
    a = ms_date[2] + ms_date[1] + ms_date[0]  #Check-4
    #print(a)            #20231130
    mid_date = date(int(ms_date[2]), int(ms_date[1]), int(ms_date[0]))

    dayy = mid_date.strftime('%A')
    #print(dayy)       #Saturday
    hol[a] = dayy
    # sss = int(ms_date[0])+1
    # print(sss)      #29
    #ms_date[0] = str(sss)
    incremented_date = mid_date + timedelta(days=1)
    mid_date = incremented_date
    year = str(mid_date.year)
    month = str(mid_date.month)
    day = str(mid_date.day)
    final_day = ""
    if (len(day) == 1):
      final_day = "0" + day
    else:
      final_day = day
    final_month = ""
    if (len(month) == 1):
      final_month = "0" + month
    else:
      final_month = month
    ms_date = [final_day, final_month, year]
    #mid_date = date(int(ms_date[2]),int(ms_date[1]),int(ms_date[0]))
  #     print(mid_date)
  #     print(ms_date)
  #     print("Checkpoint-1#")
  # print("Heloo-1")
  # print(hol)


def main():
  """
    The main function orchestrating the execution of various tasks.

    The function first attempts to read the calendar and work with its data. If successful, it proceeds to read the slot data,
    first-year data, and additional file data. If any of these steps encounters an error, an appropriate flag is set.

    After handling the initial setup, the user is prompted to input the semester start date, semester end date, mid-semester start date,
    and mid-semester holy date. The dates are processed, and holidays are added to the 'hol' dictionary using the 'holiday_add' function.
    Finally, the 'do_mail' function is called with the processed dates to execute the necessary actions related to Google Calendar.
    """
  # read_calender_error = read_calender()
  # print(read_calender_error)

  error_flag = 0
  # if (read_calender_error == 0):
  #   error_flag = work_all_calender_data()
  # print(sub)
  # print("\n")
  # print(hol)
  # print(read_calender_error, " ",error_flag)

  error_flag = read_work_calender()
  read_slot_data_error = read_slotdata_file()

  # read_first_year_error_flag = read_first_year()

  #   print(read_slot_data_error)
  #   # print(read_first_year_error_flag)

  rest_file = input(
      "Please provide the location of the file which contain details all courses(list_of_Course.csv) in csv format: "
  )

  rest_file_error = reading_rest_file(rest_file)

  # print(slot_data)
  # print(not_in_slot)
  if (error_flag == 0 and read_slot_data_error == 0 and rest_file_error == 0):

    # s_date = input("please provide the semester start date (format dd/mm/yyyy):")
    # e_date = input("please provide the semester end date(format dd/mm/yyyy):")
    # mid_date = input("please provide the semester starting mid-sem-date(format dd/mm/yyyy) :")
    # holy_date = input("please provide the semester last mid_sem-holy-date(format dd/mm/yyyy):")
    s_date = takingInput("please provide the semester start date:")
    e_date = takingInput("please provide the semester end date:")
    mid_date = takingInput("please provide the starting mid-sem-date :")
    end_mid_date = takingInput("please provide the last mid_sem Exam:")
    holy_start_date = takingInput(
        "please provide the first mid_sem Break Date:")
    holy_date = takingInput("please provide the last date of mid_sem break:")

    res = correct_order(s_date, e_date, mid_date, holy_date)

    if (res == 0):
      print("Dates inputed correctly...:)")
    elif (res == 1):
      print("Last date of midsem break is wrong, Kindly re-enter!")
      sys.exit()
    elif (res == 2):
      print("Staring date of midsem exam is wrong, Kindly re-enter!")
      sys.exit()
    elif (res == 3):
      print("Ending date of semester is wrong, Kindly re-enter!")
      sys.exit()

    start_date = s_date.split("/")
    end_date = e_date.split("/")
    ms_date = mid_date.split("/")
    ms_endDate = end_mid_date.split("/")
    holys_startDate = holy_start_date.split("/")
    holys_date = holy_date.split("/")

    s_date = date(int(start_date[2]), int(start_date[1]), int(start_date[0]))
    e_date = date(int(end_date[2]), int(end_date[1]), int(end_date[0]))
    mid_date = date(int(ms_date[2]), int(ms_date[1]), int(ms_date[0]))
    mid_endDate = date(int(ms_endDate[2]), int(ms_endDate[1]),
                       int(ms_endDate[0]))
    holy_stDate = date(int(holys_startDate[2]), int(holys_startDate[1]),
                       int(holys_startDate[0]))
    holy_date = date(int(holys_date[2]), int(holys_date[1]),
                     int(holys_date[0]))

    holiday_add(ms_date, ms_endDate, mid_date, mid_endDate)
    holiday_add(holys_startDate, holys_date, holy_stDate, holy_date)
    print(hol)
    print(sub)
    do_mail(s_date, e_date)

    # print(first_year_data)


if __name__ == "__main__":
  main()

#{'202307': ' Saturday', '202308': ' Tuesday', '202309': ' Thursday', '20230907': ' Thursday', '20231002': 'Monday', '202310': ' Tuesday', '202311': ' Monday', '202312': ' Monday', '20230922': 'Friday', '2023923': 'Saturday', '2023924': 'Sunday', '2023925': 'Monday', '2023926': 'Tuesday', '2023927': 'Wednesday', '2023928': 'Thursday', '2023929': 'Friday', '2023930': 'Saturday', '20231001': 'Sunday', '20231003': 'Tuesday', '20231004': 'Wednesday', '20231005': 'Thursday', '20231006': 'Friday', '20231007': 'Saturday', '20231008': 'Sunday'}
#['202309', '20230907', '20231002', '202311', '202312', '2023925', '2023928', '20231005']

#{'202307': ' Saturday', '202308': ' Tuesday', '202309': ' Thursday', '20230907': ' Thursday', '20231002': 'Monday', '202310': ' Tuesday', '202311': ' Monday', '202312': ' Monday', '20230922': 'Friday', '2023923': 'Saturday', '2023924': 'Sunday', '2023925': 'Monday', '2023926': 'Tuesday', '2023927': 'Wednesday', '2023928': 'Thursday', '2023929': 'Friday', '2023930': 'Saturday', '20231001': 'Sunday', '20231003': 'Tuesday', '20231004': 'Wednesday', '20231005': 'Thursday', '20231006': 'Friday', '20231007': 'Saturday', '20231008': 'Sunday'}
#['202309', '20230907', '20231002', '202311', '202312', '2023925', '2023928', '20231005']

# done
# all done