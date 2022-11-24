#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#                  Copyright (c) 2022 Ward
#
#====================================================================

# Template source: https://templates.office.com/en-us/invoice-with-sales-tax-tm03986960

import xl
import std/random
randomize()

# Load an Excel template.
var workbook = xl.load("files/template.xltx")
var sheet = workbook.active

# Clear the example values.
sheet.range("B11:D20").value = ""

# Fill sample data.
for i in 1..rand(5..9):
  var rc = "B10".rc
  sheet[rc.row + i, rc.col].value = "Item #" & $i
  sheet[rc.row + i, rc.col + 2].value = rand(50..300)

# Update to formulas.
sheet["D21"].formula = "SUM(D11:D20)"
sheet["D21"].value = 0

sheet["D23"].formula = "D21*D22"
sheet["D23"].value = 0

sheet["D25"].formula = "D21+D23+D24"
sheet["D25"].value = 0

workbook.save("invoice.xlsx")
