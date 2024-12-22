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

   #make dictionary of caller, guest_list
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