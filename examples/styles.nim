#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#                  Copyright (c) 2022 Ward
#
#====================================================================

# Example source: https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-vb-net-excel-style-formatting/202

import xl

var workbook = newWorkbook()
var sheet = workbook.add("Styles")

sheet.col(0).width = 30
sheet.col(1).width = 35

sheet[0, 0].value = "Style"
sheet[0, 1].value = "Result"

sheet.row(0).style = XlStyle(
  font: XlFont(size: 15.0, bold: true, color: XlColor(rgb: "44546A")),
  border: XlBorder(bottom: XlSide(style: "thick", color: XlColor(rgb: "4472C4")))
)

var row = 0

var side = XlSide(style: "thin", color: XlColor(rgb: "FC0101"))
row.inc(2)
sheet[row, 0].value = "border"
sheet[row, 1].border = XlBorder(
  left: side, right: side, top: side, bottom: side, diagonal: side,
  diagonalUp: true, diagonalDown: true
)

row.inc(2)
sheet[row, 0].value = "fill.patternFill"
sheet[row, 1].fill = XlFill(
  patternFill: XlPattern(
    patternType: "lightGrid",
    fgColor: XlColor(rgb: "FF00B050"),
    bgColor: XlColor(rgb: "FFFFFF00")
  )
)

row.inc(2)
sheet[row, 0].value = "font.color"
sheet[row, 1].font = XlFont(color: XlColor(rgb: "0070C0"))
sheet[row, 1].value = "XlColor(rgb: \"0070C0\")"

row.inc(2)
sheet[row, 0].value = "font.italic"
sheet[row, 1].font = XlFont(italic: true)
sheet[row, 1].value = "true"

row.inc(2)
sheet[row, 0].value = "font.name"
sheet[row, 1].font = XlFont(name: "Comic Sans MS")
sheet[row, 1].value = "Comic Sans MS"

row.inc(2)
sheet[row, 0].value = "font.vertAlign"
sheet[row, 1].font = XlFont(vertAlign: "superscript")
sheet[row, 1].value = "superscript"

row.inc(2)
sheet[row, 0].value = "font.size"
sheet[row, 1].font = XlFont(size: 18.0)
sheet[row, 1].value = "18.0"

row.inc(2)
sheet[row, 0].value = "font.strike"
sheet[row, 1].font = XlFont(strike: true)
sheet[row, 1].value = "true"

row.inc(2)
sheet[row, 0].value = "font.underline"
sheet[row, 1].font = XlFont(underline: "double")
sheet[row, 1].value = "double"

row.inc(2)
sheet[row, 0].value = "font.bold"
sheet[row, 1].font = XlFont(bold: true)
sheet[row, 1].value = "true"

row.inc(2)
sheet[row, 0].value = "alignment.horizontal"
sheet[row, 1].alignment = XlAlignment(horizontal: "center")
sheet[row, 1].value = "center"

row.inc(2)
sheet[row, 0].value = "alignment.indent"
sheet[row, 1].alignment = XlAlignment(indent: 5)
sheet[row, 1].value = "five"

row.inc(2)
sheet[row, 0].value = "alignment.verticalText"
sheet[row, 1].alignment = XlAlignment(verticalText: true)
sheet[row, 1].value = "true"

row.inc(2)
sheet[row, 0].value = "numFmt.code"
sheet[row, 1].numFmt = XlNumFmt(code: "#.##0,00 [$Krakozhian Money Units]")
sheet[row, 1].value = 1234

row.inc(2)
sheet[row, 0].value = "alignment.textRotation"
sheet[row, 1].alignment = XlAlignment(textRotation: 35)
sheet[row, 1].value = "35 degrees up"

row.inc(2)
sheet[row, 0].value = "alignment.shrinkToFit"
sheet[row, 1].alignment = XlAlignment(shrinkToFit: true)
sheet[row, 1].value = "This property is set to true so this text appears shrunk."

row.inc(2)
sheet[row, 0].value = "alignment.vertical"
sheet[row, 1].alignment = XlAlignment(vertical: "top")
sheet[row, 1].value = "top"
sheet.row(row).height = 35

row.inc(2)
sheet[row, 0].value = "alignment.wrapText"
sheet[row, 1].alignment = XlAlignment(wrapText: true)
sheet[row, 1].value = "This property is set to true so this text appears broken into multiple lines."

workbook.save("styles.xlsx")
