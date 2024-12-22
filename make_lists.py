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
   print(f"{workbook.sheetnames=}")
   sheet = workbook.active
   for row in sheet.rows:
      print(f'{row[0].value} -> {row[1].value}')

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