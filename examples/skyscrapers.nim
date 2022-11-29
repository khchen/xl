#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-create-write-excel-file/402

import xl

# Create new empty workbook and add new sheet.
var workbook = newWorkbook()
var sheet = workbook.add("Skyscrapers")

# Write title to Excel cell.
sheet["A1"].value = "List of tallest buildings (2021):"

# Sample data for writing into an Excel file.
const titles = ["Rank", "Building", "City", "Country", "Metric", "Imperial", "Floors", "Built (Year)"]
const skyscrapers = [
  (1, "Burj Khalifa", "Dubai", "United Arab Emirates", 828.0, 2717, 163, 2010),
  (2, "Shanghai Tower", "Shanghai", "China", 632.0, 2073, 128, 2015),
  (3, "Abraj Al-Bait Clock Tower", "Mecca", "Saudi Arabia", 601.0, 1971, 120, 2012),
  (4, "Ping An Finance Centre", "Shenzhen", "China", 599.0, 1965, 115, 2017),
  (5, "Lotte World Tower", "Seoul", "South Korea", 554.5, 1819, 123, 2016),
  (6, "One World Trade Center", "New York City", "United States", 541.3, 1776, 104, 2014),
  (7, "Guangzhou CTF Finance Centre", "Guangzhou", "China", 530.0, 1739, 111, 2016),
  (7, "Tianjin CTF Finance Centre", "Tianjin", "China", 530.0, 1739, 98, 2019),
  (9, "China Zun", "Beijing", "China", 528.0, 1732, 108, 2018),
  (10, "Taipei 101", "Taipei", "Taiwan", 508.0, 1667, 101, 2004),
  (11, "Shanghai World Financial Center", "Shanghai", "China", 492.0, 1614, 101, 2008),
  (12, "International Commerce Centre", "Hong Kong", "China", 484.0, 1588, 118, 2010),
  (13, "Central Park Tower", "New York City", "United States", 472.0, 1550, 98, 2020),
  (14, "Lakhta Center", "St. Petersburg", "Russia", 462.0, 1516, 86, 2019),
  (15, "Landmark 81", "Ho Chi Minh City", "Vietnam", 461.2, 1513, 81, 2018),
  (16, "Changsha IFS Tower T1", "Changsha", "China", 452.1, 1483, 88, 2018),
  (17, "Petronas Tower 1", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998),
  (17, "Petronas Tower 2", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998),
  (19, "Zifeng Tower", "Nanjing", "China", 450.0, 1476, 89, 2010),
  (19, "Suzhou IFS", "Suzhou", "China", 450.0, 1476, 98, 2019)
]

# Set row formatting.
sheet.row("1").style = XlStyle(
  font: XlFont(name: "Calibri", size: 15.0, bold: true, color: XlColor(rgb: "44546A")),
  border: XlBorder(bottom: XlSide(style: "thick", color: XlColor(rgb: "4472C4")))
)

# Set columns width.
sheet.col("A").width = 8  # Rank
sheet.col("B").width = 30 # Building
sheet.col("C").width = 16 # City
sheet.col("D").width = 20 # Country
sheet.col("E").width = 9  # Metric
sheet.col("F").width = 11 # Imperial
sheet.col("G").width = 9  # Floors
sheet.col("H").width = 9  # Built (Year)
sheet.col("I").width = 4  # Top 10
sheet.col("J").width = 5  # Top 20

# Write header data to Excel cells.
sheet{"A3:H3"}.value = titles
sheet{"E3:F3"}.move("E4")
sheet["E3"].value = "Height"

sheet{"A3:A4"}.merge() # Rank
sheet{"B3:B4"}.merge() # Building
sheet{"C3:C4"}.merge() # City
sheet{"D3:D4"}.merge() # Country
sheet{"E3:F3"}.merge() # Height
sheet{"G3:G4"}.merge() # Floors
sheet{"H3:H4"}.merge() # Built (Year)

template solidPattern(color: string): untyped =
  XlPattern(patternType: "solid", fgColor: XlColor(rgb: color))

# Set header cells formatting.
let header = sheet{"A3:H4"}
header.alignment = XlAlignment(horizontal: "center", vertical: "center", wrapText: true)
header.fill = XlFill(patternFill: solidPattern("ED7D31"))
header.font = XlFont(name: "Calibri", size: 11.0, bold: true, color: XlColor(rgb: "FFFFFF"))

# Write "Top 10" cells.
var style: XlStyle
style.alignment = XlAlignment(horizontal: "center", vertical: "center")
style.font = XlFont(name: "Calibri", size: 11.0, bold: true)

let top10 = sheet{"I5:I14"}
top10.value = "T o p   1 0"
top10.merge()

style.alignment.textRotation = -90
style.fill.patternFill = solidPattern("C6EFCE")
top10.style = style

# Write "Top 20" cells.
let top20 = sheet{"J5:J24"}
top20.value = "T o p   2 0"
top20.merge()
style.alignment.verticalText = true
style.fill.patternFill = solidPattern("FFEB9C")
top20.style = style

sheet{"I15:I24"}.merge()
sheet{"I15:I24"}.style = style

# Write sample data and formatting to Excel cells.
for i in 0..<skyscrapers.len:
  var row = sheet{(i + 4, 0), (i + 4, 7)}
  row.value = skyscrapers[i]

  if i mod 2 == 0:
    row.fill = XlFill(patternFill: solidPattern("DDEBF7"))

  for cell in row:
    if cell.rc.col == 0:
      cell.alignment = XlAlignment(horizontal: "center")

    if cell.rc.col > 3:
      cell.font = XlFont(name: "Courier New", size: 11.0)
    else:
      cell.font = XlFont(name: "Calibri", size: 11.0)

    if cell.rc.col == 4:
      cell.numFmt = XlNumFmt(code: "#\" m\"")

    if cell.rc.col == 5:
      cell.numFmt = XlNumFmt(code: "#\" ft\"")

sheet{"A5:H24"}.outline = XlSide(style: "medium")
sheet{"A5:H24"}.vertical = XlSide(style: "thin")
sheet{"H5:H24"}.right = XlSide(style: "thin")
sheet{"A5:J24"}.outline = XlSide(style: "medium")
sheet{"A5:I14"}.outline = XlSide(style: "medium")

workbook.save("skyscrapers.xlsx")
