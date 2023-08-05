#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

import xl
import std/[times, strformat]

proc writer() =
  # Create an empty workbook and add a new sheet.
  var workbook = newWorkbook()
  var sheet = workbook.add("Demo")

  # Write data to a cell.
  sheet.cell("A1").value = "Lettter"
  sheet.cell("B1").value = "Number"
  sheet["C1"].value = "Formula" # cell syntactic sugar `[]`.
  sheet["D1"].value = "Date"

  # Using range and data array.
  sheet.range("A2:A4").value = ["A", "B", "C", "D", "E"]
  sheet{"B2:B4"}.value = [1, 2, 3, 4, 5] # range syntactic sugar `{}`.

  # Using range and data tuple.
  sheet{"C2:D4"}.value = (
    0, now(),
    0, now() + initDuration(days=1),
    0, now() + initDuration(days=2),
  )

  # Using turned range and data tuple.
  sheet{"C5:D6"}.turn.value = (
    0, 0,
    now() + initDuration(days=3), now() + initDuration(days=4),
  )

  # Write formula to a cell.
  sheet["C2"].formula = "POWER(B2, 2)"

  # Write formulas to cells in a range.
  sheet{"C3:C6"}.formula = [
    "POWER(B3, 2)",
    "POWER(B4, 2)",
    "POWER(B5, 2)",
    "POWER(B6, 2)",
  ]

  # Add more data via range iterator.
  for cell in sheet{"A5:A11"}:
    cell.value = $chr(cell.rc.row - 1 + 'A'.ord)

  # Add more data via cell proc for column object.
  let colB = sheet.col("B")
  for i in 4..10:
    colB.cell(i).value = i

  # Add more formulas.
  let colC = sheet.col("C")
  for i in 6..10:
    colC[i].formula = &"POWER(B{i + 1}, 2)"

  # Add more dates.
  const D = "D1".rc.col # get col value of "D" column
  for i in 6..10:
    sheet[i, D].date = now() + initDuration(days=(i-1))

  # Set style and width of column.
  sheet.col("D").numFmt = XlNumFmt(code: "yyyy/m/d")
  sheet.col("D").width = 15

  # Set alignment of row.
  sheet.row("1").alignment = XlAlignment(horizontal: "center")

  # Set border of range.
  sheet.range.outline = XlSide(style: "thick", color: XlColor(rgb: "00B050"))
  sheet.range.horizontal = XlSide(style: "thin", color: XlColor(rgb: "00B050"))
  sheet.range.vertical = XlSide(style: "thin", color: XlColor(rgb: "00B050"))

  # Set style of range.
  sheet.range.font = XlFont(name: "Arial")

  # Set style of collection.
  var coll = sheet.collection(empty=true)
  coll.add "A1:D1"
  coll.font = XlFont(name: "Arial", bold: true)
  coll.bottom = XlSide(style: "medium", color: XlColor(rgb: "00B050"))

  # save the file
  workbook.save("demo.xlsx")

proc reader() =
  # Load workbook and get the active worksheet.
  var workbook = xl.load("demo.xlsx")
  var sheet = workbook.active

  assert sheet.count == 44
  assert sheet.dimension.name == "A1:D11"

  # Iterate over rows in range and then cells in row.
  var values: seq[string]
  for row in sheet{"A2:B4"}.rows:
    for cell in row:
      values.add cell.value

  assert values == @["A", "1", "B", "2", "C", "3"]

  # Iterate over columns in range and then cells in column.
  values = @[]
  for col in sheet{"A2:B4"}.cols:
    for cell in col:
      values.add cell.value

  assert values == @["A", "B", "C", "1", "2", "3"]

  # Iterate over cells in range. Row first by defualt.
  values = @[]
  for cell in sheet{"A2:B4"}:
    values.add cell.value

  assert values == @["A", "1", "B", "2", "C", "3"]

  # Iterate over cells in turned range. Column first.
  values = @[]
  for cell in sheet{"A2:B4"}.turn:
    values.add cell.value

  assert values == @["A", "B", "C", "1", "2", "3"]

  # Get styles of a cell.
  assert sheet["A1"].font == XlFont(name: "Arial", bold: true)
  assert sheet["A1"].alignment == XlAlignment(horizontal: "center")

  # Get styles of of a row.
  assert sheet.row(0).alignment == XlAlignment(horizontal: "center")

  # `$` of styles can copy and paste as nim code.
  assert $sheet["A1"].font == """XlFont(name: "Arial", bold: true)"""
  assert $sheet.row(0).alignment == """XlAlignment(horizontal: "center")"""

when isMainModule:
  writer()
  reader()
