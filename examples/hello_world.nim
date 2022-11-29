#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-vb-net-excel-library/601

import xl

var workbook = newWorkbook()
var sheet = workbook.add("Hello World")

# The standard way to get cell object.
sheet.cell("A1").value = "English:"
sheet.cell(0, 1).value = "Hello"

# `[]` is the syntactic sugar of cell().
sheet["A2"].value = "Russian:"
sheet[1, 1].value = "Здравствуйте"

# Using range object instead of cell object.
sheet.range("A3:A3").value = "Chinese:"
sheet.range("B3", (2, 1)).value = "你好"

# Write array into range.
sheet.range("A4:B4").value = ["Japanese:", "こんにちは"]

# `{}` is the syntactic sugar of range().
sheet{"A6:J6"}.value = "In order to see Russian, Chinese, and Japanese characters you need to have appropriate fonts on your PC."
sheet{(5, 0), "J6"}.merge()

workbook.save("hello_world.xlsx")
