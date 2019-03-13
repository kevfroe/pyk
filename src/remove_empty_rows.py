# by krowe
# TODO:
#   - read text files
#   - convert text files to .xlsx files

import optparse
import openpyxl
import os

DEBUG = False

def row_is_empty(row):
  for cell in row:
    if cell.value != None:
      return False
  return True

def get_row_number(row):
  return row[1].row

def print_row(row):
  print("  row {}: {}".format(get_row_number(row), ",".join([str(cell.value) if cell.value != None else "" for cell in row])))

def remove_empty_rows(opts):
  wb = openpyxl.load_workbook(opts.file_in)

  for sheetname in wb.sheetnames:
    rows_to_delete = []
    sheet = wb[sheetname]
    
    if DEBUG:
      print("SHEET: {}".format(sheetname))

    for row in sheet.iter_rows():
      if DEBUG:
        print_row(row)

      if row_is_empty(row):
        rows_to_delete.append(get_row_number(row))
    
    if DEBUG:
      print("rows to delete: {}".format(rows_to_delete))

    rows_to_delete = sorted(rows_to_delete, reverse=True)
    for row_num in rows_to_delete:
      sheet.delete_rows(row_num)

    if DEBUG:
      for row in sheet.iter_rows():
        print_row(row)

  wb.save(filename=opts.file_out)

def main():
  parser = optparse.OptionParser()
  parser.add_option("--file-in",  dest="file_in",  help="input *.xlsx FILE",  metavar="FILE_IN")
  parser.add_option("--file-out", dest="file_out", help="output *.xlsx FILE", metavar="FILE_OUT")

  (opts, _) = parser.parse_args()

  remove_empty_rows(opts)
  
if __name__ == "__main__":
  main()