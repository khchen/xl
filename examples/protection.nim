#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://xlsxwriter.readthedocs.io/example_protection.html

import xl

var workbook = newWorkbook()
var sheet = workbook.add("Sheet1")

# Format the columns to make the text more visible.
sheet.col(0).width = 40

# Write a locked, unlocked and hidden cell.
sheet["A1"].value = "Cell B1 is locked. It cannot be edited."
sheet["A2"].value = "Cell B2 is unlocked. It can be edited."
sheet["A3"].value = "Cell B3 is hidden. The formula isn't visible."

sheet.range("B1:B3").formula = "1+2"

# Create some cell formats with protection properties.
sheet["B1"].protection = XlProtection() # Locked by default.
sheet["B2"].protection = XlProtection(locked: false)
sheet["B3"].protection = XlProtection(hidden: true)

# Turn worksheet protection on.
sheet.protect()

workbook.save("protection.xlsx")
