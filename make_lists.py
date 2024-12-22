#!/usr/bin/env python3

import sys
try:
   from openpyxl import load_workbook
   import PyPDF4
except Exception as e:
   print(e)
   sys.exit(1)

import argparse
from pathlib import Path

def make_assignments(in_filename):

   workbook = load_workbook(in_filename, data_only=True)
   sheetnames = workbook.sheetnames
   # print(f"{workbook.sheetnames=}")

   EXPECTED_SHEETNAMES = ['guest-to-caller', 'callers', 'guests']
   if sheetnames != EXPECTED_SHEETNAMES:
      print("Error: expected '{EXPECTED_SHEETNAMES}' sheet names; found '{workbook.sheetnames}' in file '{in_filename}'")
      return False

   #make dictionary of caller, [guests]
   mapping_dict = {}
   is_header = True
   for row in workbook['callers'].rows:
      # skip header row; there is only 1 column: the list of callers
      if is_header:
         is_header = False
      else:
         mapping_dict[row[0].value] = []
   
   is_header = True
   for row in workbook['guest-to-caller'].rows:
      # columns: 0=Guest, 1=Caller, 2=Note
      if is_header:
         is_header = False
      else:
         # add the guest name and note to the caller's list:
         mapping_dict[row[1].value].append([row[0].value, row[2].value])
         # print(f'{row[0].value} -> {row[1].value}')
   # print(f"{mapping_dict=}\n")
   '''
   mapping_dict={'Caroline': [['Guest1', 'new regular']], 'Tina': [], 'Peter': [], 'Rebecca': [['Guest2', 'Substitute this week only']], 'Maria': [], 'Barb': [], 'Lisa': [['Guest3', None]], 'Do-Not-Call': []}
   '''

   # make a dictionary of guest data to be used for generating reports.
   guest_dict = {}
   is_header = True
   for row in workbook['guests'].rows:
      # columns: 0=First, 1=Last, 2=UserName, 3=Password, 4=Town, 5=Phone, 6=Notes
      if is_header:
         is_header = False
      else:
         guest_dict[row[2].value]= {'First':row[0].value, 'Last':row[1].value, 'PW':row[3].value,
                                 'Town':row[4].value, 'Phone':row[5].value, 'Notes':row[6].value}      
   # print(f"{guest_dict=}")
   '''
   guest_dict={'Guest1': {'First': 'Guest', 'Last': 1.0, 'PW': 'secret', 'Town': 'Newbury', 'Phone': '978.555.0000', 'Notes': 'call early'}, 'Guest2': {'First': 'Guest', 'Last': 2.0, 'PW': 'secret', 'Town': 'Newbury', 'Phone': '978.555.0000', 'Notes': 'call 3 times'}, 'Guest3': {'First': 'Guest', 'Last': 3.0, 'PW': 'secret', 'Town': 'Newbury', 'Phone': '978.555.0000', 'Notes': 'call late'}}
   '''

   callers_with_no_guests = []
   for caller, guests in mapping_dict.items():
      if len(guests) == 0:
         callers_with_no_guests.append(caller)
      else:
         print(f"{caller}:") # {guests=}")
         for guest_id_note in guests:
            this_guest= guest_dict[guest_id_note[0]]
            if guest_id_note[1] is not None:
               this_weeks_guest_note = guest_id_note[1]
            else:
               this_weeks_guest_note = ''
            print(f"\t{this_guest['First']}\t{this_guest['Last']}\t{guest_id_note[0]}\
\t{this_guest['PW']}\t{this_weeks_guest_note}\t{this_guest['Town']}\t{this_guest['Phone']}\t{this_guest['Notes']}")
   
   print(f"These callers have no guests: {callers_with_no_guests}")

   return True


if __name__ == "__main__":
   argParser = argparse.ArgumentParser()
   argParser.add_argument("input", type=str, help="input filename with path")

   args = argParser.parse_args()
   # print(f'\n\t{args.input= } {args.auth_path= }\n\t{args.dont_email= } {args.parse_only= }')

   if args.input is None:
      sys.exit("No file selected to parse.")
   elif Path(args.input).is_file():
      filename = args.input
   else:
      sys.exit("file selected is not a file.")

   rc = make_assignments(filename)

   if rc:
      print(f"generated report")
      