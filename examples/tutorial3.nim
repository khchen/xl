#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://xlsxwriter.readthedocs.io/tutorial03.html

import xl
import std/times

# Some data we want to write to the worksheet.
const expenses = [
    ("Rent", "2013-01-13", 1000),
    ("Gas", "2013-01-14", 100),
    ("Food", "2013-01-16", 300),
    ("Gym", "2013-01-20", 50),
  ]

 # Create a workbook and add a worksheet.
var workbook = newWorkbook()
var sheet = workbook.add("Tutorial3")

# Create empty collections.
var bolds = sheet.collection()
var moneys = sheet.collection()
var dates = sheet.collection()

# Adjust the column width.
sheet.col(1).width = 15

# Write some data headers.
sheet["A1"].value = "Item"
sheet["B1"].value = "Date"
sheet["C1"].value = "Cost"

# `{}` is syntactic sugar of range.
bolds.add sheet{"A1:C1"}

# Start from the first cell below the headers.
var (row, col) = "A2".rc

# Iterate over the data and write it out row by row.
for (item, date, cost) in expenses:
  sheet[row, col].value = item
  sheet[row, col + 1].value = parse(date, "yyyy-MM-dd")
  sheet[row, col + 2].value = cost
  dates.add (row, col + 1)
  moneys.add (row, col + 2)
  row.inc

# Write a total using a formula (without prefix "=").
sheet[row, 0].value = "Total"
sheet[row, 2].formula = "SUM(C2:C5)"
bolds.add (row, 0)
moneys.add (row, 2)

# Set styles to cells in collections.
bolds.font = XlFont(bold: true)
moneys.numFmt = XlNumFmt(code: "$#,##0")
dates.numFmt = XlNumFmt(code: "mmmm d yyyy")

workbook.save("tutorial3.xlsx")
