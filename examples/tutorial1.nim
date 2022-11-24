#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#                  Copyright (c) 2022 Ward
#
#====================================================================

# Example source: https://xlsxwriter.readthedocs.io/tutorial01.html

import xl

# Some data we want to write to the worksheet.
const expenses = {
  "Rent": 1000,
  "Gas": 100,
  "Food": 300,
  "Gym": 50,
}

# Create a workbook and add a worksheet.
var workbook = newWorkbook()
var sheet = workbook.add("Tutorial1")

# Start from the first cell. Rows and columns are zero indexed.
var row = 0

# Iterate over the data and write it out row by row.
for (item, cost) in expenses:
  sheet.cell(row, 0).value = item
  sheet.cell(row, 1).value = cost
  row.inc

# Write a total using a formula (without prefix "=").
sheet.cell(row, 0).value = "Total"
sheet.cell(row, 1).formula = "SUM(B1:B4)"

workbook.save("tutorial1.xlsx")
