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

More examples: https://github.com/khchen/xl/tree/master/examples.

## Performance

Not bad....

Create or load workbook with `experimental=true` or use `cells` iterator for workbook to get best performance.

## Docs
* https://khchen.github.io/xl

## License
Read license.txt for more details.

Copyright (c) 2022 Kai-Hung Chen, Ward. All rights reserved.

## Donate
If this project help you reduce time to develop, you can give me a cup of coffee :)

[![paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://paypal.me/khchen0915?country.x=TW&locale.x=zh_TW)
