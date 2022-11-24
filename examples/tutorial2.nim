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
var sheet = workbook.add("Tutorial2")

# Create empty collections.
var bold = sheet.collection()
var money = sheet.collection()

# Write some data headers. `[]` is syntactic sugar of cell in sheet.
sheet["A1"].value = "Item"
sheet["B1"].value = "Cost"
bold.add sheet.range("A1:B1")

# Start from the first cell below the headers.
var row = 1

# Iterate over the data and write it out row by row.
for (item, cost) in expenses:
  sheet[row, 0].value = item
  sheet[row, 1].value = cost
  money.add (row, 1)
  row.inc

# Write a total using a formula (without prefix "=").
sheet[row, 0].value = "Total"
sheet[row, 1].formula = "SUM(B2:B5)"
bold.add (row, 0)
money.add (row, 1)

# Set styles to cells in collections.
bold.font = XlFont(bold: true)
money.numFmt = XlNumFmt(code: "$#,##0")

workbook.save("tutorial2.xlsx")
