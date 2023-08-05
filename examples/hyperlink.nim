#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://xlsxwriter.readthedocs.io/example_hyperlink.html

import xl

const RedFont = XlFont(
  color: XlColor(rgb: "ff0000"),
  bold: true,
  underline: "single",
  size: 12.0
)

var workbook = newWorkbook()
var sheet = workbook.add("Hyperlink")

sheet.col(0).width = 30

sheet["A1"].hyperlink = "https://nim-lang.org/"
sheet["A1"].value = "Nim Homepage"
sheet["A1"].font = RedFont

sheet["A3"].hyperlink = "https://www.google.com/"
sheet["A3"].value = "Google"
sheet["A3"].font = RedFont

workbook.save("hyperlink.xlsx")
