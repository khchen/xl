#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://github.com/GemBoxLtd/GemBox.Spreadsheet.Examples/tree/master/C%23/Common%20Uses/Reading

import xl
import std/strutils

# Function to output a cell.
proc output(cell: XlCell) =
  var value = cell.value
  if value == "": value = "EMPTY"
  if value.len >= 9: value = value[0..<9] & "..."
  if not cell.isNumber: value = '"' & value & '"'
  stdout.write(alignLeft(value, 15))

# Method1: using cells iterator.
proc reader1(workbook: XlWorkbook) =
  for name in workbook.sheetNames:
    echo "Worksheet: ", name

    # This is the fastest way to iterate over readonly cells of a sheet.
    # Notice: the empty cell may be skipped.
    var lastRow = 0
    for cell in workbook.cells(name):
      if cell.rc.row != lastRow:
        lastRow = cell.rc.row
        stdout.write("\n")

      cell.output()

    stdout.write("\n")
  stdout.write("\n")

# Method2: using standard way.
proc reader2(workbook: XlWorkbook) =
  for name in workbook.sheetNames:
    echo "Worksheet: ", name

    # The standard way to iterate rows and editable cells.
    for row in workbook.sheet(name).range.rows:
      for cell in row:
        cell.output()

      stdout.write("\n")
    stdout.write("\n")
  stdout.write("\n")

when isMainModule:
  var workbook = xl.load("files/data.xlsx")
  workbook.reader1()
  workbook.reader2()
