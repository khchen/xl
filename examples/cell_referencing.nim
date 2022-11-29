#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#                  Copyright (c) 2022 Ward
#
#====================================================================

# Example source: https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-excel-range/204

import xl
import std/strformat

var workbook = newWorkbook()
var sheet = workbook.add("Referencing")

# Adjust the column width.
for i in 0..<1024:
  sheet.col(i).width = 10

# Referencing cells from sheet using cell names and indexes.
sheet.cell("A1").value = "Cell A1."
sheet.cell(1, 0).value = "Cell in 2nd row and 1st column [A2]."

# Referencing cells from row using cell names and indexes.
sheet.row("4").cell("B").value = "Cell in row 4 and column B [B4].";
sheet.row(4).cell(1).value = "Cell in 5th row and 2nd column [B5].";

# Referencing cells from column using cell names and indexes.
sheet.col("C").cell("7").value = "Cell in column C and row 7 [C7].";
sheet.col(2).cell(7).value = "Cell in 3rd column and 8th row [C8].";

# Referencing cell range using A1 notation [G2:N12].
var r = sheet.range("G2:N12")
r[0].value = &"From {r.first.name} to {r.last.name}"
r[1, 0].value = &"From ({r.first.row}, {r.first.col}) to ({r.last.row}, {r.last.col})"
r.outline = XlSide(style: "thick", color: XlColor(rgb: "FF0000"))

# Referencing cell range using absolute position [I5:M11].
r = sheet.range((4, 8), (10, 12))
r[0].value = &"From {r.first.name} to {r.last.name}"
r[1, 0].value = &"From ({r.first.row}, {r.first.col}) to ({r.last.row}, {r.last.col})"
r.outline = XlSide(style: "medium", color: XlColor(rgb: "00B050"))

# Referencing cell range using relative position [K8:L9].
r = sheet.range("K8".rc, "L9".rc)
r[0].value = &"From {r.first.name} to {r.last.name}"
r[1, 0].value = &"From ({r.first.row}, {r.first.col}) to ({r.last.row}, {r.last.col})"
r.outline = XlSide(style: "thin", color: XlColor(rgb: "0070C0"))

workbook.save("cell_referencing.xlsx")