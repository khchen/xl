#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://xlsxwriter.readthedocs.io/example_merge_rich.html

import xl

var workbook = newWorkbook()
var sheet = workbook.add "Merge Rich String"

sheet.range("B2:E5").merge()

sheet["B2"].riches = {
  "This is ": XlFont(),
  "red": XlFont(color: XlColor(rgb: "FF0000")),
  " and this is ": XlFont(),
  "blue": XlFont(color: XlColor(rgb: "0000FF"))
}

sheet["B2"].alignment = XlAlignment(
  horizontal: "center",
  vertical: "center"
)

workbook.save("merge_rich_string.xlsx")
