[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://paypal.me/khchen0915?country.x=TW&locale.x=zh_TW)

# Xl
Xl is pure nim library to create, read, and modify open XML spreadsheet (Excel) files.

## Features
- Pure nim, only dependency is [zippy](https://github.com/guzba/zippy "zippy") (pure nim too).
- Support .xlsx and .xltx. format.
- Read and write string, number, date, formula, hyperlink, rich string, and styles of cells, ranges, and collections.
- Move range, copy range, merge/unmerge range, insert rows, insert columns, delete rows, delete columns.
- Row height and style, column width and style.
- All styles are supported (number format, font, fill, border, alignment, and protection).

## Examples
An hello world program:
```nim
import xl

var workbook = newWorkbook()
var sheet = workbook.add("Hello World")

sheet.cell("A1").value = "English:"
sheet.cell(0, 1).value = "Hello"

workbook.save("hello_world.xlsx")
```

Load, modify, and write back to xlsx:
```nim
import xl

var workbook = xl.load("filename.xlsx")
var sheet = workbook.active

# Read the value of a cell.
echo sheet.cell("A1").value

# Read the number of a cell
echo sheet.cell("A1").number

# Modify style and value.
sheet.cell("A1").value = "Hello"
sheet.cell("A1").font = XlFont(size: 14.0, bold: true)

# Read the style of a cell, the output is vaild nim object construction expression.
echo sheet.cell("A1").style # all style, include number format, font, fill, etc.
echo sheet.cell("A1").font # font style only

assert $sheet.cell("A1").font == """XlFont(size: 14.0, bold: true)"""

# Reset to default style.
sheet.cell("A2").style = default(XlStyle)

# After modifing, the workbook can be save into another file.
workbook.save("another.xlsx")
```

More examples: https://github.com/khchen/xl/tree/main/examples.

## Performance

Not bad....

Create or load workbook with `experimental=true` or use [cells](https://khchen.github.io/xl/#cells.i%2CXlWorkbook%2Cint "cells") iterator for workbook to get best performance.

The following code can parse [large.xlsx](https://github.com/theorchard/openpyxl/blob/master/openpyxl/benchmarks/files/large.xlsx "large.xlsx") less than 5 seconds on my computer.

```nim
import xl

var workbook = xl.load("large.xlsx", experimental=true)
var total = 0.0

for cell in workbook.cells("Sheet1"):
  if cell.isNumber:
    total += cell.number
```

```
# Nim Compiler Version 1.6.8
nim r -d:release -d:danger --opt:speed -d:lto --gc:arc large.nim
```



## Docs
* https://khchen.github.io/xl

## License
Read license.txt for more details.

Copyright (c) 2022 Kai-Hung Chen, Ward. All rights reserved.

## Donate
If this project help you reduce time to develop, you can give me a cup of coffee :)

[![paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://paypal.me/khchen0915?country.x=TW&locale.x=zh_TW)
