#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

## Open XML Spreadsheet (Excel) Library for Nim.

# TODO
#  - freeze panes
#  - cell selection
#  - image? comment? chart?

import xl/private/xn
import zippy/ziparchives
import std/[strutils, os, tables, options, strscans, strtabs, algorithm,
  critbits, sets, hashes, streams, times, math]

template `/`(x: string): string =
  '/' & x

template `/`(x, y: string): string =
  x & '/' & y

template `/../`(head, tail: string): string =
  os.`/../`(head, tail).replace("\\", "/")

const
  NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  NsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types"
  NsRelationships = "http://schemas.openxmlformats.org/package/2006/relationships"
  NsDocumentRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  NsCoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  NsExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  NsDc = "http://purl.org/dc/elements/1.1/"
  NsDcTerms = "http://purl.org/dc/terms/"
  NsXSI = "http://www.w3.org/2001/XMLSchema-instance"

  TypeHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  TypeWorksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  TypeStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  TypeSharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  TypeOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  TypeCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
  TypeExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"

  MimeXml = "application/xml"
  MimeWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  MimeStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
  MimeSharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
  MimeRelationships = "application/vnd.openxmlformats-package.relationships+xml"
  MimeCoreProperties = "application/vnd.openxmlformats-package.core-properties+xml"
  MimeExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
  MimeSheetMain = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
  MimeTemplateMain = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
  MimeCalcChain = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"

  Xl = "xl"
  DocPropsApp = "docProps/app.xml"
  DocPropsCore = "docProps/core.xml"
  Workbook = Xl/"workbook.xml"
  SharedStrings = Xl/"sharedStrings.xml"
  Styles = Xl/"styles.xml"
  ContentTypes = "[Content_Types].xml"

  BuiltinFormats = [
    "General", "0", "0.00", "#,##0", "#,##0.00", "($#,##0_);($#,##0)",
    "($#,##0_);[Red]($#,##0)", "($#,##0.00_);($#,##0.00)", "($#,##0.00_);[Red]($#,##0.00)",
    "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??", "m/d/yy", "d-mmm-yy", "d-mmm", "mmm-yy",
    "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yy h:mm", "General", "General",
    "General", "General", "General", "General", "General", "General", "General", "General",
    "General", "General", "General", "General", "(#,##0_);(#,##0)", "(#,##0_);[Red](#,##0)",
    "(#,##0.00_);(#,##0.00)", "(#,##0.00_);[Red](#,##0.00)",
    "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
    "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
    "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
    "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    "mm:ss", "[h]:mm:ss", "mm:ss.0", "##0.0E+0", "@"
  ]

type
  XlCellType = enum
    ctNumber = "n"
    ctSharedString = "s"
    ctBoolean = "b"
    ctError = "e"
    ctInlineStr = "inlineStr"
    ctStr = "str"

  XlError* = object of CatchableError
    ## Default error object.

  XlRC* = tuple[row: int, col: int]
    ## Represent a cell position.

  XlRCRC* = tuple[first: XlRC, last: XlRC]
    ## Represent a range.

  XlObject = ref object
    raw: string
    xn: XlNode
    postProcess: (string, string)

  XlChange = ref object
    numFmt: Option[XlNumFmt]
    font: Option[XlFont]
    fill: Option[XlFill]
    border: Option[XlBorder]
    alignment: Option[XlAlignment]
    protection: Option[XlProtection]

  XlSharedStrings = object
    xSharedStrings: XlNode
    table: OrderedTable[string, tuple[index: int, isNew: bool]]
    newCount: int

  XlSharedStyles = object
    xStyles: XlNode
    styles: OrderedTable[tuple[oldStyle: int, change: XlChange], int]
    numFmts: Table[XlNumFmt, int]
    fonts: Table[XlFont, int]
    fills: Table[XlFill, int]
    borders: Table[Xlborder, int]
    xfCount: int

  XlRel = object
    typ: string
    target: string
    external: bool

  XlType = object
    typ: string
    override: bool

  XlRels = Table[string, XlRel]

  XlTypes = Table[string, XlType]

  XlProperties* = object
    ## Represent workbook properties.
    title*: string
    subject*: string
    creator*: string
    keywords*: string
    description*: string
    lastModifiedBy*: string
    created*: string
    modified*: string
    category*: string
    contentStatus*: string
    application*: string
    manager*: string
    company*: string

  XlSheetProtection* = object
    ## Represent sheet protection properties.
    sheet*: Option[bool]
    objects*: Option[bool]
    scenarios*: Option[bool]
    selectLockedCells*: Option[bool]
    selectUnlockedCells*: Option[bool]
    formatCells*: Option[bool]
    formatColumns*: Option[bool]
    formatRows*: Option[bool]
    insertColumns*: Option[bool]
    insertRows*: Option[bool]
    insertHyperlinks*: Option[bool]
    deleteColumns*: Option[bool]
    deleteRows*: Option[bool]
    sort*: Option[bool]
    autoFilter*: Option[bool]
    pivotTables*: Option[bool]
    password*: Option[string]

  XlWorkbook* = ref object
    ## Represent a workbook.
    contents: Table[string, XlObject]
    sheets: seq[XlSheet]
    sharedStrings: XlSharedStrings
    sharedStyles: XlSharedStyles
    rels: XlRels
    types: XlTypes
    properties: XlProperties
    active: int
    experimental: bool

  XlSheet* = ref object
    ## Represent a worksheet.
    workbook: XlWorkbook
    obj: XlObject
    name: string
    hidden: bool
    color: XlColor
    protection: XlSheetProtection
    cells: Table[XlRC, XlCell]
    rows: Table[int, XlRow]
    cols: Table[int, XlCol]
    merges: HashSet[XlRCRC]
    rels: XlRels
    first: XlRC
    last: XlRC

  XlRange* = ref object
    ## Represent a rectangle range.
    sheet: XlSheet
    first: XlRC
    last: XlRC
    turned: bool

  XlCollection* = ref object
    ## Represent a collection of cells.
    sheet: XlSheet
    rcs: HashSet[XlRC]

  XlStylable = ref object of RootRef
    sheet: XlSheet
    change: XlChange
    style: int

  XlRow* = ref object of XlStylable
    ## Represent a row of a sheet.
    row: int
    height: float
    hidden: bool
    outlineLevel: int
    collapsed: bool
    thickTop: bool
    thickBot: bool
    ph: bool

  XlCol* = ref object of XlStylable
    ## Represent a column of a sheet.
    col: int
    width: float
    hidden: bool

  XlCell* = ref object of XlStylable
    ## Represent a cell of a sheet.
    rc: XlRC
    formula: string
    value: string
    hyperlink: string
    riches: XlRiches
    ct: XlCellType
    readonly: bool

  XlNumFmt* = object
    ## Represent number format.
    code*: Option[string]

  XlColor* = object
    ## Represent a color.
    ##
    ## Possible values of rgb:
    ## 6 or 8 hexdigits presents RGB or ARGB color.
    ##
    ## Possible values of indexed:
    ## `0`..`64` system default color index.
    ##
    ## Possible values of tint:
    ## `-1.0`..`1.0`.
    `auto`*: Option[bool]
    rgb*: Option[string]
    theme*: Option[int]
    indexed*: Option[int]
    tint*: Option[float]

  XlFont* = object
    ## Represent font style.
    ##
    ## Possible values of family:
    ## `0`..`14`.
    ##
    ## Possible values of underline:
    ## `""`, `"single"`, `"double"`, `"singleAccounting"`, `"doubleAccounting"`.
    ##
    ## Possible values of vertAlign:
    ## `"superscript"`, `"subscript"`, `"baseline"`.
    ##
    ## Possible values of scheme:
    ## `"major"`, `"minor"`.
    name*: Option[string]
    charset*: Option[int]
    family*: Option[int]
    size*: Option[float]
    bold*: Option[bool]
    italic*: Option[bool]
    strike*: Option[bool]
    underline*: Option[string]
    vertAlign*: Option[string]
    color*: Option[XlColor]
    scheme*: Option[string]

  XlGradient* = object
    ## Represent gradient fill style.
    ##
    ## Possible values of gradientType:
    ## `"linear"`, `"path"`.
    gradientType*: Option[string]
    degree*: Option[float]
    left*, right*, top*, bottom*: Option[float]
    stops*: seq[XlColor]

  XlPattern* = object
    ## Represent pattern fill style.
    ##
    ## Possible values of patternType:
    ## `"none"`, `"solid"`, `"mediumGray"`, `"darkGray"`, `"lightGray"`,
    ## `"darkHorizontal"`, `"darkVertical"`, `"darkDown"`, `"darkUp"`,
    ## `"darkGrid"`, `"darkTrellis"`, `"lightHorizontal"`, `"lightVertical"`,
    ## `"lightDown"`, `"lightUp"`, `"lightGrid"`, `"lightTrellis"`, `"gray125"`,
    ## `"gray0625"`.
    patternType*: Option[string]
    fgColor*: Option[XlColor]
    bgColor*: Option[XlColor]

  XlFill* = object
    ## Represent fill style.
    patternFill*: Option[XlPattern]
    gradientFill*: Option[XlGradient]

  XlAlignment* = object
    ## Represent alignment style.
    ##
    ## Possible values of horizontal:
    ## `"general"`, `"left"`, `"center"`, `"right"`, `"fill"`, `"justify"`,
    ## `"centerContinuous"`, `"distributed"`.
    ##
    ## Possible values of vertical:
    ## `"top"`, `"center"`, `"bottom"`, `"justify"`, `"distributed"`.
    ##
    ## Possible values of textRotation:
    ## `-90`..`180` or `255` (vertical text).
    horizontal*: Option[string]
    vertical*: Option[string]
    textRotation*: Option[int]
    wrapText*: Option[bool]
    shrinkToFit*: Option[bool]
    indent*: Option[int]
    relativeIndent*: Option[int]
    justifyLastLine*: Option[bool]
    readingOrder*: Option[int]
    verticalText*: Option[bool]

  XlProtection* = object
    ## Represent protection setting.
    locked*: Option[bool]
    hidden*: Option[bool]

  XlSide* = object
    ## Represent a side of border.
    ##
    ## Possible values of style:
    ## `"dashDot"`, `"dashDotDot"`, `"dashed"`, `"dotted"`, `"double"`,
    ## `"hair"`, `"medium"`, `"mediumDashDot"`, `"mediumDashDotDot"`,
    ## `"mediumDashed"`, `"slantDashDot"`, `"thick"`, `"thin"`.
    style*: Option[string]
    color*: Option[XlColor]

  XlBorder* = object
    ## Represent border style.
    outline*: Option[bool]
    diagonalUp*: Option[bool]
    diagonalDown*: Option[bool]
    left*: Option[XlSide]
    right*: Option[XlSide]
    top*: Option[XlSide]
    bottom*: Option[XlSide]
    diagonal*: Option[XlSide]
    vertical*: Option[XlSide]
    horizontal*: Option[XlSide]

  XlStyle* = object
    ## Represent all styles.
    numFmt*: XlNumFmt
    font*: XlFont
    fill*: XlFill
    border*: XlBorder
    alignment*: XlAlignment
    protection*: XlProtection

  XlRich = tuple
    text: string
    font: XlFont

  XlRiches* = seq[XlRich]
    ## Represent rich string styles.

const
  EmptyRowCol = XlRC (-1, -1)

converter toSomeOption*[T: bool|int|float|string|XlColor|XlSide|XlPattern|XlGradient](x: T): Option[T] {.inline.} =
  ## Syntactic sugar for option type to avoid trivial `some`.
  some x

# for XlChange ref type in hashtable
proc `==`(x, y: XlChange): bool =
  if x.isNil and y.isNil:
    return true
  elif (x.isNil and not y.isNil) or (y.isNil and not x.isNil):
    return false
  else:
    return x[] == y[]

# for XlChange ref type in hashtable
proc hash(x: XlChange): Hash {.inline.} =
  if x.isNil:
    return hash(cast[pointer](x))
  else:
    return hash(x[])

iterator sortedItems[A](t: HashSet[A]): A =
  var s = newSeqOfCap[A](t.len)
  for i in t: s.add i
  s.sort()
  for i in s: yield i

iterator sortedKeys[A, B](t: Table[A, B]): lent A =
  var keys: seq[A]
  for k in t.keys: keys.add k
  keys.sort()
  for k in keys: yield k

iterator sortedValues[A, B](t: var Table[A, B]): lent B =
  # var keys = newSeqOfCap[A](t.len)
  # for k in t.keys: keys.add k
  # keys.sort()
  # for k in keys: yield t[k]

  var s = newSeqOfCap[(A, ptr B)](t.len)
  for k, v in t.mpairs:
    s.add (k, addr v)

  s.sort() do (a, b: (A, ptr B)) -> int:
    system.cmp(a[0], b[0])

  for p in s:
    yield p[1][]

iterator sortedPairs[A, B](t: var Table[A, B]): (A, var B) =
  var s = newSeqOfCap[(A, ptr B)](t.len)
  for k, v in t.mpairs:
    s.add (k, addr v)

  s.sort() do (a, b: (A, ptr B)) -> int:
    system.cmp(a[0], b[0])

  for p in s:
    yield (p[0], p[1][])

template hasBoolOptionAttr(x: XlNode, attr: XlNsName|string): Option[bool] =
  let val = x[attr]
  case val
  of "true":
    some(true)
  of "false":
    some(false)
  of "":
    none(bool)
  else:
    try:
      some((bool parseInt(val)))
    except:
      none(bool)

proc `[]=`[T](x: XlNode, key: string, opt: Option[T]) {.inline.} =
  if opt.isSome:
    when T is string|SomeNumber:
      x[key] = $opt.get

    elif T is bool:
      x[key] = $opt.get.int

    else:
      {.error: "unknow type " & $T.}

proc expand(x: XlNode, ns: string, order: openarray[string]) =
  var
    cursor = 0
    index = 0

  while cursor < order.len:
    if index >= x.count:
      discard x.addChildXlNode(ns\order[cursor])

    elif x[index].tag.ns != ns: # skip tag for other namespace
      index.inc
      continue

    elif x[index].tag != ns\order[cursor]:
      discard x.addChildXlNode(ns\order[cursor], index)

    cursor.inc
    index.inc

proc shrink(x: XlNode, ns: string, preserves: openarray[string]) =
  for i in countdown(x.count - 1, 0):
    let child = x[i]
    var isDelete = true
    if child.count == 0 and (child.attrs == nil or child.attrs.len == 0):
      for tag in preserves:
        if child.tag == ns\tag:
          isDelete = false
          break

      if isDelete:
        x.delete(i)

proc updateCount(x: XlNode) {.inline.} =
  x["count"] = $x.count

proc parse(x: XlObject) =
  if x.raw != "" and x.xn.isNil:
    x.xn = parseXn(x.raw)
    x.raw = ""

proc newXlObject(x: sink XlNode): XlObject {.inline.} =
  result = XlObject(xn: x)

proc newXlObject(raw: sink string): XlObject {.inline.} =
  result = XlObject(raw: raw)

proc rc*(name: string): XlRC =
  ## Convert cell reference (e.g. "A1") to XlRC tuple.
  let name = name.toUpperAscii
  var col, row, idx = 0
  if scanp(name, idx,
      +{'A'..'Z'} -> (col = col * 26 + ord($_) - ord('A') + 1),
      +{'0'..'9'} -> (row = row * 10 + ord($_) - ord('0'))):
    result.row = row
    result.col = col

  result.row.dec
  result.col.dec
  if result.row < 0 or result.col < 0:
    raise newException(XlError, "invalid cell name")

proc rc*(xc: XlCell): XlRC {.inline.} =
  ## Return XlRC tuple of a cell.
  result = xc.rc

proc rcrc*(name: string): XlRCRC =
  ## Convert range reference (e.g. "A1:C3") to XlRCRC tuple.
  let sp = name.split(':', maxsplit=1)
  if sp.len == 2:
    result[0] = sp[0].rc
    result[1] = sp[1].rc
  else:
    result[0] = name.rc
    result[1] = result[0]

proc rcrc*(xr: XlRange): XlRCRC {.inline.} =
  ## Return XlRCRC tuple of a XlRange object.
  result = (xr.first, xr.last)

proc name*(rc: XlRC): string =
  ## Convert a XlRC tuple to named reference (e.g. "A1").
  if rc.row < 0 or rc.col < 0:
    raise newException(XlError, "invalid XlRowCol")

  var col = rc.col
  while col >= 0:
    result.insert $(chr('A'.ord + col mod 26))
    col = col div 26 - 1

  result.add $(rc.row + 1)

proc name*(rcrc: XlRCRC): string {.inline.} =
  ## Convert a XlRCRC tuple to range reference (e.g. "A1:C3").
  result = rcrc[0].name
  result.add ':'
  result.add rcrc[1].name

proc name*(xc: XlCell): string {.inline.} =
  ## Return named reference of a cell.
  result = xc.rc.name

proc value*(x: XlRow): int {.inline.} =
  ## Return row index value of a XlRow object.
  result = x.row

proc value*(x: XlCol): int {.inline.} =
  ## Return col index value of a XlCol object.
  result = x.col

proc len*(xs: XlSheet): int {.inline.} =
  ## Return total cells count of a sheet.
  result = xs.cells.len

proc count*(xs: XlSheet): int {.inline.} =
  ## Return total cells count of a sheet.
  result = xs.cells.len

proc isEmpty*(xr: XlRange): bool {.inline.} =
  ## Check if a XlRange object is empty or not.
  result = xr.first.row < 0 or xr.first.col < 0 or
    xr.last.row < 0 or xr.last.col < 0

proc isEmpty*(xs: XlSheet): bool {.inline.} =
  ## Check if a sheet is empty or not.
  result = xs.count == 0

proc `$`(x: XlObject): string =
  if x.xn != nil:
    result.add xlHeader
    result.add $$x.xn
    if x.postProcess != ("", ""):
      result = result.replace(x.postProcess[0], x.postProcess[1])
  else:
    result = x.raw

proc `$`*(x: XlWorkbook): string =
  ## Convert XlWorkbook object to string.
  result = "XlWorkBook("
  result.add "sheets: ["
  for s in x.sheets:
    result.addQuoted s.name
    result.add ", "
  result.removeSuffix(", ")
  result.add "])"

proc `$`*(x: XlSheet): string =
  ## Convert XlSheet object to string.
  result = "XlSheet("
  result.add "name: "
  result.addQuoted x.name
  result.add ", count: "
  result.addQuoted x.count
  if not x.isEmpty:
    result.add ", dimension: "
    result.addQuoted (x.first, x.last).name
  result.add ")"

proc `$`*(x: XlCell|XlRow|XlCol): string =
  ## Convert XlCell, XlRow, or XlCol object to string.
  result = $x.type & "("
  when x is XlCell:
    result.add x.rc.name
  elif x is XlRow:
    result.add $(x.row + 1)
  elif x is XlCol:
    result.add (0, x.col).name
    result.removeSuffix("1")
  result.add ")"

proc `$`*(x: XlRange): string =
  ## Convert XlRange object to string.
  result = "XlRange("
  if not x.isEmpty:
    result.add (x.first, x.last).name
  result.add ")"

proc `$`*(x: XlCollection): string =
  ## Convert XlCollection object to string.
  result = "XlCollection("
  for rc in x.rcs.sortedItems:
    result.add rc.name
    result.add ","
  result.removeSuffix(",")
  result.add ")"

proc `$`*(x: XlNumFmt|XlColor|XlFont|XlAlignment|XlBorder|XlSide|XlFill|
    XlPattern|XlProtection|XlSheetProtection): string =
  ## Convert style related object to string.
  ## The output can copy and paste as nim code.
  result = $x.type & "("
  for k, v in fieldPairs(x):
    if v.isSome:
      result.add k & ": "
      result.addQuoted v.get
      result.add ", "
  result.removeSuffix(", ")
  result.add ")"

proc `$`*(x: seq[Option[XlColor]]): string =
  ## Convert style related object to string.
  ## The output can copy and paste as nim code.
  result.add "@["
  for i in x:
    if i.isSome:
      result.add $i.get
      result.add ", "
  result.removeSuffix(", ")
  result.add "]"

proc `$`*(x: XlGradient): string =
  ## Convert style related object to string.
  ## The output can copy and paste as nim code.
  result = "XlGradient("
  for k, v in fieldPairs(x):
    when v is Option:
      if v.isSome:
        result.add k & ": "
        result.addQuoted v.get
        result.add ", "
    elif v is seq:
      result.add k & ": "
      result.add $v
      result.add ", "

  result.removeSuffix(", ")
  result.add ")"

proc `$`*(x: XlStyle|XlProperties): string =
  ## Convert style related object to string.
  ## The output can copy and paste as nim code.
  result = $x.type & "("
  for k, v in fieldPairs(x):
    if v != default(v.type):
      result.add k & ": "
      result.addQuoted v
      result.add ", "
  result.removeSuffix(", ")
  result.add ")"

proc `$`(x: openArray[XlRich]): string =
  result = "{"
  for rich in x:
    result.addQuoted rich.text
    result.add ": "
    result.add $rich.font
    result.add ", "
  if x.len == 0:
    result.add ":"
  else:
    result.removeSuffix(", ")
  result.add "}"

proc `$`*(x: XlRiches): string =
  ## Convert XlRiches to string.
  ## The output can copy and paste as nim code.
  result = $toOpenArray(x, 0, x.high)

proc `$`*[T](x: array[T, XlRich]): string =
  ## Convert XlRiches to string.
  ## The output can copy and paste as nim code.
  result = $toOpenArray(x, 0, x.high)

proc isParsed(sheet: XlSheet): bool {.inline.} =
  return not sheet.obj.xn.isNil

proc `{}`(xw: XlWorkbook, file: string): XlNode =
  if file in xw.contents:
    let obj = xw.contents[file]
    obj.parse()
    return obj.xn
  else:
    raise newException(XlError, file & " not found")

proc dup(cell: XlCell): XlCell =
  result = XlCell()
  result[] = cell[]
  if result.change != nil: # change is ref, deep copy it
    result.change = XlChange()
    result.change[] = cell.change[]

template updateFirstLast(x: XlSheet|XlRange, rc: XlRC) =
  x.first.row = min(x.first.row, rc.row)
  x.first.col = min(x.first.col, rc.col)
  x.last.row = max(x.last.row, rc.row)
  x.last.col = max(x.last.col, rc.col)

proc update(xs: XlSheet, rc: XlRC) =
  if xs.isEmpty:
    xs.first = rc
    xs.last = rc
  else:
    xs.updateFirstLast(rc)

proc update(xs: XlSheet, c: XlCell) =
  let rc = c.rc
  xs.update(rc)
  xs.cells[rc] = c

proc update(xs: XlSheet) =
  if xs.isEmpty:
    xs.first = EmptyRowCol
    xs.last = EmptyRowCol
  else:
    xs.first = (int.high, int.high)
    xs.last = (int.low, int.low)
    for rc in xs.cells.keys:
      xs.updateFirstLast(rc)

template check(obj: untyped, field: untyped, error: bool, assertion: untyped) =
  if obj.field.isSome:
    var value {.inject.} = obj.field.get
    if not assertion:
      if error:
        raise newException(XlError, "Invalid " & $obj)
      else:
        obj.field = none obj.field.get.type

template checkValidate(obj: untyped, field: untyped, error: bool) =
  if obj.field.isSome:
    validate(obj.field.get, error)
    if obj.field.get == default(obj.field.get.type):
      obj.field = none obj.field.get.type

proc validate(color: var XlColor, error: bool) =
  check(color, rgb, error): value.len in {6, 8} and value.allCharsInSet(HexDigits)
  check(color, theme, error): value >= 0
  check(color, indexed, error): value in 0..64
  check(color, tint, error): value in -1.0..1.0

proc validate(side: var XlSide, error: bool) =
  check(side, style, error):
    value in ["dashDot","dashDotDot", "dashed","dotted", "double", "hair",
      "medium", "mediumDashDot", "mediumDashDotDot", "mediumDashed",
      "slantDashDot", "thick", "thin"]

  checkValidate(side, color, error)

proc validate(numFmt: var XlNumFmt, error: bool) =
  check(numFmt, code, error): value != ""

proc validate(font: var XlFont, error: bool) =
  check(font, charset, error): value >= 0
  check(font, family, error): value in 0..14
  check(font, size, error): value >= 0
  check(font, underline, error):
    value in ["", "single", "double", "singleAccounting", "doubleAccounting"]

  check(font, vertAlign, error):
    value in ["superscript", "subscript", "baseline"]

  check(font, scheme, error):
    value in ["major", "minor"]

  checkValidate(font, color, error)

proc validate(gradient: var XlGradient, error: bool) =
  check(gradient, gradientType, error):
    value in ["linear", "path"]

  for i in countdown(gradient.stops.high, 0):
    validate(gradient.stops[i], error)
    if gradient.stops[i] == default(XlColor):
      gradient.stops.delete(i)

proc validate(pattern: var XlPattern, error: bool) =
  check(pattern, patternType, error):
    value in ["none", "solid", "mediumGray", "darkGray", "lightGray",
      "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid",
      "darkTrellis", "lightHorizontal", "lightVertical", "lightDown",
      "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"]

  checkValidate(pattern, fgColor, error)
  checkValidate(pattern, bgColor, error)

proc validate(fill: var XlFill, error: bool) =
  checkValidate(fill, patternFill, error)
  checkValidate(fill, gradientFill, error)

proc validate(border: var XlBorder, error: bool) =
  checkValidate(border, left, error)
  checkValidate(border, right, error)
  checkValidate(border, top, error)
  checkValidate(border, bottom, error)
  checkValidate(border, diagonal, error)
  checkValidate(border, vertical, error)
  checkValidate(border, horizontal, error)

proc validate(align: var XlAlignment, error: bool) =
  check(align, horizontal, error):
    value in ["general", "left", "center", "right", "fill", "justify",
      "centerContinuous", "distributed"]

  check(align, vertical, error):
    value in ["top", "center", "bottom", "justify", "distributed"]

  check(align, textRotation, error): value in -90..180 or value == 255
  check(align, indent, error): value >= 0
  check(align, relativeIndent, error): value >= 0
  check(align, readingOrder, error): value >= 0

proc validate(align: var XlProtection, error: bool) {.inline.} =
  # nothing to do for XlProtection
  discard

proc option[T](x: XlNode, name: XlNsName): Option[T] {.inline.} =
  let val = x[name]
  if val == "":
    return none T

  try:
    when T is string: result = some val
    elif T is SomeInteger: result = some(T parseInt(val))
    elif T is SomeFloat: result = some(T parseFloat(val))
    elif T is bool: result = some(val == "true" or parseInt(val).bool)
    else:
      {.error: "unknow type " & $T.}

  except:
    return none T

proc childOption[T](x: XlNode, name: XlNsName, attr: XlNsName): Option[T] {.inline.} =
  var child = x.child(name)
  if child.isNil:
    return none T

  return option[T](child, attr)

proc parseColor(xColor: XlNode): XlColor =
  result.`auto` = option[bool](xColor, NsMain\"auto")
  result.rgb = option[string](xColor, NsMain\"rgb")
  result.indexed = option[int](xColor, NsMain\"indexed")
  result.theme = option[int](xColor, NsMain\"theme")
  result.tint = option[float](xColor, NsMain\"tint")
  validate(result, error=false)

proc newChildColor(x: XlNode, name: XlNsName, color: XlColor): XlNode =
  result = newXlNode(x, name)
  result["auto"] = color.`auto`
  result["rgb"] = color.rgb
  result["indexed"] = color.indexed
  result["theme"] = color.theme
  result["tint"] = color.tint

proc parseAlignment(xStyles: XlNode, style: int): XlAlignment =
  if xStyles.hasChildN(NsMain\"cellXfs", style, xf) and
      xf.hasChild(NsMain\"alignment", xAlign):
    result.horizontal = option[string](xAlign, NsMain\"horizontal")
    result.vertical = option[string](xAlign, NsMain\"vertical")
    result.textRotation = option[int](xAlign, NsMain\"textRotation")
    result.wrapText = option[bool](xAlign, NsMain\"wrapText")
    result.shrinkToFit = option[bool](xAlign, NsMain\"shrinkToFit")
    result.indent = option[int](xAlign, NsMain\"indent")
    result.relativeIndent = option[int](xAlign, NsMain\"relativeIndent")
    result.justifyLastLine = option[bool](xAlign, NsMain\"justifyLastLine")
    result.readingOrder = option[int](xAlign, NsMain\"readingOrder")
    if result.textRotation == 255:
      result.verticalText = true
    validate(result, error=false)

proc newChildAlignment(x: XlNode, name: XlNsName, align: XlAlignment): XlNode =
  result = newXlNode(x, name)
  result["horizontal"] = align.horizontal
  result["vertical"] = align.vertical
  result["wrapText"] = align.wrapText
  result["shrinkToFit"] = align.shrinkToFit
  result["indent"] = align.indent
  result["relativeIndent"] = align.relativeIndent
  result["justifyLastLine"] = align.justifyLastLine
  result["readingOrder"] = align.readingOrder

  if align.verticalText.isSome and align.verticalText.get:
    result["textRotation"] = 255
  elif align.textRotation.isSome:
    if align.textRotation.get < 0: # -90 ~ -1 should be 180 ~ 91
      result["textRotation"] = abs(align.textRotation.get) + 90
    else:
      result["textRotation"] = align.textRotation

proc parseProtection(xStyles: XlNode, style: int): XlProtection =
  if xStyles.hasChildN(NsMain\"cellXfs", style, xf) and
      xf.hasChild(NsMain\"protection", xProtection):
    result.locked = option[bool](xProtection, NsMain\"locked")
    result.hidden = option[bool](xProtection, NsMain\"hidden")
    validate(result, error=false)

proc newChildProtection(x: XlNode, name: XlNsName, protection: XlProtection): XlNode =
  result = newXlNode(x, name)
  result["locked"] = protection.locked
  result["hidden"] = protection.hidden

proc childColorOption(x: XlNode, name: XlNsName): Option[XlColor] =
  var child = x.child(name)
  if child.isNil:
    return none XlColor

  result = some parseColor(child)

proc childSideOption(x: XlNode, name: XlNsName): Option[XlSide] =
  var child = x.child(name)
  if child.isNil:
    return none XlSide

  var side = XlSide()
  side.style = option[string](child, NsMain\"style")
  side.color = childColorOption(child, NsMain\"color")
  return if side == default(XlSide): none XlSide else: some side

proc addChildAttr[T](x: XlNode, name: XlNsName, attr: string, opt: Option[T]) =
  if opt.isSome:
    var child = x.addChildXlNode(name)
    if attr != "":
      child[attr] = opt

proc addChildColor(x: XlNode, name: XlNsName, color: Option[XlColor]) =
  if color.isSome:
    x.add x.newChildColor(name, color.get)

proc addChildSide(x: XlNode, name: XlNsName, side: Option[XlSide]) =
  if side.isSome:
    var xSide = x.addChildXlNode(name)
    var side = side.get
    xSide["style"] = side.style
    xSide.addChildColor(NsMain\"color", side.color)

# Both MS Office ant Libre Office don't check applyNumberFormat, applyFont, applyFill,
# applyAlignment, applyBorder of cellXfs.

proc parseNumFmt(xStyles: XlNode, style: int): XlNumFmt =
  if xStyles.hasChildN(NsMain\"cellXfs", style, xf) and
      # xf.hasTrueAttr(NsMain\"applyNumberFormat") and
      xf.hasIntAttr(NsMain\"numFmtId", fmtId):

    if fmtId < 50:
      result.code = some BuiltinFormats[fmtId]

    elif fmtId < 164:
      result.code = some "General"

    elif xStyles.hasChild(NsMain\"numFmts", numFmts):
      for i in numFmts.find(NsMain\"numFmt"):
        if i.hasIntAttr(NsMain\"numFmtId", n) and fmtId == n:
          result.code = some i[NsMain\"formatCode"]

    validate(result, error=false)

proc applyNumFmt(xStyles: XlNode, numFmt: XlNumFmt): int =
  if numFmt.code.isNone: return -1 # should not happen?
  let fmt = numFmt.code.get

  for i, f in BuiltinFormats:
    if fmt == f: return i

  xStyles.editChild(NsMain\"numFmts", xNumFmts, 0):
    var id = 0
    for xNumFmt in xNumFmts.find(NsMain\"numFmt"):
      try: id = max(parseInt(xNumFmt["numFmtId"]), id)
      except ValueError: discard

    id = max(id + 1, 164)
    discard xNumFmts.addChildXlNode(NsMain\"numFmt", -1,
      {"numFmtId": $id, "formatCode": fmt}
    )
    xNumFmts.updateCount()
    return id

proc parseFont(xStyles: XlNode, style: int): XlFont =
  if xStyles.hasChildN(NsMain\"cellXfs", style, xf) and
      # xf.hasTrueAttr(NsMain\"applyFont") and
      xf.hasIntAttr(NsMain\"fontId", fontId) and
      xStyles.hasChildN(NsMain\"fonts", fontId, xFont):

    result.name = childOption[string](xFont, NsMain\"name", NsMain\"val")
    result.charset = childOption[int](xFont, NsMain\"charset", NsMain\"val")
    result.family = childOption[int](xFont, NsMain\"family", NsMain\"val")
    result.size = childOption[float](xFont, NsMain\"sz", NsMain\"val")
    result.bold = if xFont.hasChild(NsMain\"b"): some true else: none bool
    result.italic = if xFont.hasChild(NsMain\"i"): some true else: none bool
    result.strike = if xFont.hasChild(NsMain\"strike"): some true else: none bool
    result.vertAlign = childOption[string](xFont, NsMain\"vertAlign", NsMain\"val")
    result.color = childColorOption(xFont, NsMain\"color")
    result.scheme = childOption[string](xFont, NsMain\"scheme", NsMain\"val")

    # single underline only has a <u /> child, convert to some "single"
    if xFont.hasChild(NsMain\"u", u):
      if u.hasAttr(NsMain\"val", val):
        result.underline = some val
      else:
        result.underline = some "single"

    validate(result, error=false)

proc applyFont(xStyles: XlNode, font: XlFont): int =
  xStyles.editChild(NsMain\"fonts", xFonts, 0):
    let xFont = xFonts.addChildXlNode(NsMain\"font")
    xFonts.updateCount()
    result = xFonts.count - 1

    xFont.addChildAttr(NsMain\"name", "val", font.name)
    xFont.addChildAttr(NsMain\"charset", "val", font.charset)
    xFont.addChildAttr(NsMain\"family", "val", font.family)
    xFont.addChildAttr(NsMain\"sz", "val", font.size)
    xFont.addChildAttr(NsMain\"b", "", font.bold)
    xFont.addChildAttr(NsMain\"i", "", font.italic)
    xFont.addChildAttr(NsMain\"strike", "", font.strike)
    xFont.addChildAttr(NsMain\"vertAlign", "val", font.vertAlign)
    xFont.addChildColor(NsMain\"color", font.color)
    xFont.addChildAttr(NsMain\"scheme", "val", font.scheme)

    if font.underline.isSome:
      if font.underline.get in ["", "single"]:
        xFont.addChildAttr(NsMain\"u", "", font.underline)
      else:
        xFont.addChildAttr(NsMain\"u", "val", font.underline)

proc parseFill(xStyles: XlNode, style: int): XlFill =
  if xStyles.hasChildN(NsMain\"cellXfs", style, xf) and
      # xf.hasTrueAttr(NsMain\"applyFill") and
      xf.hasIntAttr(NsMain\"fillId", fillId) and
      xStyles.hasChildN(NsMain\"fills", fillId, xFill):

    if xFill.hasChild(NsMain\"patternFill", xPattern):
      var patternFill = XlPattern()
      patternFill.patternType = option[string](xPattern, NsMain\"patternType")
      patternFill.fgColor = childColorOption(xPattern, NsMain\"fgColor")
      patternFill.bgColor = childColorOption(xPattern, NsMain\"bgColor")
      result.patternFill = some patternFill

    if xFill.hasChild(NsMain\"gradientFill", xGradient):
      var gradientFill = XlGradient()
      gradientFill.gradientType = option[string](xGradient, NsMain\"type")
      gradientFill.degree = option[float](xGradient, NsMain\"degree")
      gradientFill.left = option[float](xGradient, NsMain\"left")
      gradientFill.right = option[float](xGradient, NsMain\"right")
      gradientFill.top = option[float](xGradient, NsMain\"top")
      gradientFill.bottom = option[float](xGradient, NsMain\"bottom")
      for xStop in xGradient.find(NsMain\"stop"):
        if xStop.hasChild(NsMain\"color", color):
          gradientFill.stops.add parseColor(color)
      result.gradientFill = some gradientFill

    validate(result, error=false)

proc applyFill(xStyles: XlNode, fill: XlFill): int =
  xStyles.editChild(NsMain\"fills", xFills, 0):
    let xFill = xFills.addChildXlNode(NsMain\"fill")
    xFills.updateCount()
    result = xFills.count - 1

    if fill.patternFill.isSome:
      var patternFill = fill.patternFill.get
      var xPattern = xFill.addChildXlNode(NsMain\"patternFill")
      xPattern["patternType"] = patternFill.patternType
      xPattern.addChildColor(NsMain\"fgColor", patternFill.fgColor)
      xPattern.addChildColor(NsMain\"bgColor", patternFill.bgColor)

    if fill.gradientFill.isSome:
      var gradientFill = fill.gradientFill.get
      var xGradient = xFill.addChildXlNode(NsMain\"gradientFill")
      xGradient["type"] = gradientFill.gradientType
      xGradient["degree"] = gradientFill.degree
      xGradient["left"] = gradientFill.left
      xGradient["right"] = gradientFill.right
      xGradient["top"] = gradientFill.top
      xGradient["bottom"] = gradientFill.bottom
      for i, color in gradientFill.stops:
        var xStop = xGradient.addChildXlNode(NsMain\"stop", -1, {"position": $i})
        xStop.addChildColor(NsMain\"color", color)

proc parseBorder(xStyles: XlNode, style: int): XlBorder =
  if xStyles.hasChildN(NsMain\"cellXfs", style, xf) and
      # xf.hasTrueAttr(NsMain\"applyBorder") and
      xf.hasIntAttr(NsMain\"borderId", borderId) and
      xStyles.hasChildN(NsMain\"borders", borderId, xBorder):

    result.outline = option[bool](xBorder, NsMain\"outline")
    result.diagonalUp = option[bool](xBorder, NsMain\"diagonalUp")
    result.diagonalDown = option[bool](xBorder, NsMain\"diagonalDown")
    result.left = childSideOption(xBorder, NsMain\"left")
    result.right = childSideOption(xBorder, NsMain\"right")
    result.top = childSideOption(xBorder, NsMain\"top")
    result.bottom = childSideOption(xBorder, NsMain\"bottom")
    result.diagonal = childSideOption(xBorder, NsMain\"diagonal")
    result.vertical = childSideOption(xBorder, NsMain\"vertical")
    result.horizontal = childSideOption(xBorder, NsMain\"horizontal")

    validate(result, error=false)

proc applyBorder(xStyles: XlNode, border: XlBorder): int =
  xStyles.editChild(NsMain\"borders", xBorders, 0):
    let xBorder = xBorders.addChildXlNode(NsMain\"border")
    xBorders.updateCount()
    result = xBorders.count - 1

    xBorder["outline"] = border.outline
    xBorder["diagonalUp"] = border.diagonalUp
    xBorder["diagonalDown"] = border.diagonalDown
    xBorder.addChildSide(NsMain\"left", border.left)
    xBorder.addChildSide(NsMain\"right", border.right)
    xBorder.addChildSide(NsMain\"top", border.top)
    xBorder.addChildSide(NsMain\"bottom", border.bottom)
    xBorder.addChildSide(NsMain\"diagonal", border.diagonal)
    xBorder.addChildSide(NsMain\"vertical", border.vertical)
    xBorder.addChildSide(NsMain\"horizontal", border.horizontal)

proc parseRpr(xRpr: XlNode): XlFont =
  result.name = childOption[string](xRpr, NsMain\"rFont", NsMain\"val")
  result.charset = childOption[int](xRpr, NsMain\"charset", NsMain\"val")
  result.family = childOption[int](xRpr, NsMain\"family", NsMain\"val")
  result.size = childOption[float](xRpr, NsMain\"sz", NsMain\"val")
  result.bold = if xRpr.hasChild(NsMain\"b"): some true else: none bool
  result.italic = if xRpr.hasChild(NsMain\"i"): some true else: none bool
  result.strike = if xRpr.hasChild(NsMain\"strike"): some true else: none bool
  result.vertAlign = childOption[string](xRpr, NsMain\"vertAlign", NsMain\"val")
  result.color = childColorOption(xRpr, NsMain\"color")
  result.scheme = childOption[string](xRpr, NsMain\"scheme", NsMain\"val")

  # single underline only has a <u /> child, convert to some "single"
  if xRpr.hasChild(NsMain\"u", u):
    if u.hasAttr(NsMain\"val", val):
      result.underline = some val
    else:
      result.underline = some "single"

  validate(result, error=false)

proc applyRpr(x: XlNode, font: XlFont) =
  var xRpr = x.addChildXlNode(NsMain\"rPr")
  xRpr.addChildAttr(NsMain\"rFont", "val", font.name)
  xRpr.addChildAttr(NsMain\"charset", "val", font.charset)
  xRpr.addChildAttr(NsMain\"family", "val", font.family)
  xRpr.addChildAttr(NsMain\"sz", "val", font.size)
  xRpr.addChildAttr(NsMain\"b", "", font.bold)
  xRpr.addChildAttr(NsMain\"i", "", font.italic)
  xRpr.addChildAttr(NsMain\"strike", "", font.strike)
  xRpr.addChildAttr(NsMain\"vertAlign", "val", font.vertAlign)
  xRpr.addChildColor(NsMain\"color", font.color)
  xRpr.addChildAttr(NsMain\"scheme", "val", font.scheme)

  if font.underline.isSome:
    if font.underline.get in ["", "single"]:
      xRpr.addChildAttr(NsMain\"u", "", font.underline)
    else:
      xRpr.addChildAttr(NsMain\"u", "val", font.underline)

proc value(riches: XlRiches): string =
  for rich in riches:
    result.add rich.text

proc save(riches: XlRiches, x: XlNode, tag: XlNsName): XlNode =
  result = x.newXlNode(tag)
  for rich in riches:
    var xr = result.addChildXlNode(NsMain\"r")

    if rich.font != default(XlFont):
      xr.applyRpr(rich.font)

    if rich.text != "":
      var xt = xr.addChildXlNode(NsMain\"t", rich.text)
      xt["xml:space"] = "preserve"

proc load(x: var XlRiches, xn: XlNode) =
  for r in xn.find(NsMain\"r"):
    var rich: XlRich
    if r.hasChild(NsMain\"t", xt):
      rich.text = xt.innerText

    if r.hasChild(NsMain\"rPr", xrpr):
      rich.font = parseRpr(xrpr)

    x.add rich

proc save(s: XlSharedStrings): XlNode =
  result = s.xSharedStrings.dup()
  result.addNameSpace(NsMain) # in case not exist

  for text, (index, isNew) in s.table:
    if isNew:
      var xsi = result.addChildXlNode(NsMain\"si")
      discard xsi.addChildXlNode(NsMain\"t", text)

proc reset(s: var XlSharedStrings) {.inline.} =
  s.newCount = 0
  s.table.clear()

proc getRiches(s: XlSharedStrings, index: int): XlRiches {.inline.} =
  if index < s.xSharedStrings.count:
    result.load(s.xSharedStrings[index])

proc get(s: XlSharedStrings, index: int): string {.inline.} =
  if index < s.xSharedStrings.count:
    return s.xSharedStrings[index].innerText

proc add(s: var XlSharedStrings, text: string): int =
  let tup = s.table.getOrDefault(text, (-1, false))
  if tup.index >= 0:
    return tup.index

  let index = s.xSharedStrings.count + s.newCount
  s.table[text] = (index, true)
  s.newCount.inc
  return index

proc load(s: var XlSharedStrings, x: XlNode) =
  s.reset()
  s.xSharedStrings = x

  for i, xsi in x.childrenPair:
    if xsi.tag != NsMain\"si": # should not happen, invaild file?
      continue

    if xsi.hasChild(NsMain\"t", xt):
      s.table[xt.innerText] = (i, false)

proc save(s: var XlSharedStyles): XlNode =

  proc addXf(xCellXfs: XlNode, oldStyle: int): XlNode =
    if oldStyle >= 0 and oldStyle < xCellXfs.count:
      result = xCellXfs[oldStyle].dup()
      xCellXfs.add result

    else:
      result = xCellXfs.addChildXlNode(NsMain\"xf", -1,
        {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0"})

  proc apply[T](xStyles: XlNode, xf: XlNode, option: Option[T],
      tab: var Table[T, int], applyFn: proc(a: XlNode, b: T): int,
      aid: string, aok: string) =

    # check option isSome, apply real change to xml by applyFn
    # cache change in table, and write id record into xml
    if option.isSome:
      let o = option.get
      if o == default(T):
        xf[aid] = "0"
        xf[aok] = "0"

      else:
        var id = tab.getOrDefault(o, -1)
        if id < 0:
          id = applyFn(xStyles, o)
          tab[o] = id

        xf[aid] = $id
        xf[aok] = "1"

  template applyChild[T](xf: XlNode, sym: untyped, newfn: T, aok: string, index: int): untyped =
    if change.sym.isSome:
      if change.sym.get == default(change.sym.get.type):
        xf.delete(NsMain\astToStr(sym))
        xf.deleteAttr(aok)
      else:
        xf.replace(xf.newfn(NsMain\astToStr(sym), change.sym.get), index)
        xf[aok] = "1"

  var xStyles = s.xStyles.dup()
  xStyles.addNameSpace(NsMain)
  xStyles.editChild(NsMain\"cellXfs", xCellXfs, -1):
    for tup, v in s.styles:
      let (oldStyle, change) = tup
      var xf = xCellXfs.addXf(oldStyle)

      apply(xStyles, xf, change.numFmt, s.numFmts, applyNumFmt, "numFmtId", "applyNumberFormat")
      apply(xStyles, xf, change.font, s.fonts, applyFont, "fontId", "applyFont")
      apply(xStyles, xf, change.fill, s.fills, applyFill, "fillId", "applyFill")
      apply(xStyles, xf, change.border, s.borders, applyBorder, "borderId", "applyBorder")

      applyChild(xf, alignment, newChildAlignment, "applyAlignment", 0)
      applyChild(xf, protection, newChildProtection, "applyProtection", -1)

    xCellXfs.updateCount()

  return xStyles

proc get(s: var XlSharedStyles, x: XlStylable): int =
  # no style changing, use the old style
  if x.change == nil:
    return x.style

  # if change is all default => reset the style
  # return -1 to indicate not to write the style record
  if x.change[] == default(XlChange()[].type):
    return -1

  # if x.style and x.change are the same
  # reuse the xf record is ok
  let styleChanged = (x.style, x.change)
  let index = s.styles.getOrDefault(styleChanged, -1)
  if index >= 0:
    return index

  # otherwise, save the style/change pair into s.styles table,
  # and returns the next index of xf record.
  # save proc will create xf records according to s.styles table.
  s.styles[styleChanged] = s.xfCount
  result = s.xfCount
  s.xfCount.inc

proc reset(s: var XlSharedStyles) =
  s.styles.clear()
  s.numFmts.clear()
  s.fonts.clear()
  s.fills.clear()
  s.borders.clear()
  if s.xStyles.hasChild(NsMain\"cellXfs", cellXfs):
    s.xfCount = cellXfs.count
  else:
    s.xfCount = 0

proc load(s: var XlSharedStyles, x: XlNode) {.inline.} =
  s.xStyles = x
  s.reset()

proc add(types: var XlTypes, key: string, typ: string, override=false) {.inline.} =
  types[key] = XlType(typ: typ, override: override)

proc add(types: var XlTypes, key: string, obj: XlType) {.inline.} =
  types[key] = obj

proc save(types: XlTypes, mime = MimeSheetMain): XlNode =
  result = newXlRootNode(NsContentTypes\"Types")
  for key, obj in types:
    var typ = obj.typ

    if obj.override:
      # overwrite the mime type of workbook
      if key.endsWith "xl/workbook.xml":
        typ = mime

      discard result.addChildXlNode(NsContentTypes\"Override", -1,
        {"PartName": key, "ContentType": typ})

    else:
      discard result.addChildXlNode(NsContentTypes\"Default", 0,
        {"Extension": key, "ContentType": typ})

proc load(types: var XlTypes, x: XlNode) =
  for child in x.children:
    if child.tag == NsContentTypes\"Default":
      types.add(child["Extension"], child["ContentType"], false)

    elif child.tag == NsContentTypes\"Override":
      types.add(child["PartName"], child["ContentType"], true)

    else:
      discard

proc getRelsFile(target: string): string =
  let sp = target.splitPath
  result = sp.head / "_rels" / sp.tail & ".rels"

proc nextId(rels: XlRels): string =
  var n = rels.len + 1
  while true:
    result = "rId" & $n
    if result notin rels: break

proc add(rels: var XlRels, rel: XlRel): string {.discardable, inline.} =
  result = rels.nextId()
  rels[result] = rel

proc add(rels: var XlRels, typ: string, target: string, external=false): string {.discardable, inline.} =
  result = rels.nextId()
  rels[result] = XlRel(typ: typ, target: target, external: external)

proc load(rels: var XlRels, x: XlNode, path: string) =
  let root = path.splitPath.head /../ ""
  for child in x.find(NsRelationships\"Relationship"):
    var rel = XlRel(
      typ: child["Type"],
      target: child["Target"],
      external: child["TargetMode"] == "External"
    )

    if not rel.external:
      # convert all path to "relative path of /"
      if rel.target.startsWith("/"):
        rel.target.removePrefix("/")

      else:
        rel.target = root / rel.target

    rels[child["Id"]] = rel

proc save(rels: var XlRels): XlNode =
  result = newXlRootNode(NsRelationships\"Relationships")
  for id, rel in rels.sortedPairs:
    var target = rel.target
    if not rel.external and target.startsWith Xl/"":
      target = /target

    discard result.addChildXlNode(NsRelationships\"Relationship", -1,
      {"Id": id, "Type": rel.typ, "Target": target})

    if rel.external:
      result[^1]["TargetMode"] = "External"

proc removeExternal(rels: var XlRels) =
  var todel: seq[string]
  for id, rel in rels:
    if rel.external:
      todel.add id

  for id in todel:
    rels.del(id)

proc load(x: var XlProperties, core: XlNode, app: XlNode) =
  if core.hasChild(NsDc\"title", xn): x.title = xn.innerText
  if core.hasChild(NsDc\"subject", xn): x.subject = xn.innerText
  if core.hasChild(NsDc\"creator", xn): x.creator = xn.innerText
  if core.hasChild(NsCoreProperties\"keywords", xn): x.keywords = xn.innerText
  if core.hasChild(NsDc\"description", xn): x.description = xn.innerText
  if core.hasChild(NsCoreProperties\"lastModifiedBy", xn): x.lastModifiedBy = xn.innerText
  if core.hasChild(NsDcTerms\"created", xn): x.created = xn.innerText
  if core.hasChild(NsDcTerms\"modified", xn): x.modified = xn.innerText
  if core.hasChild(NsCoreProperties\"category", xn): x.category = xn.innerText
  if core.hasChild(NsCoreProperties\"contentStatus", xn): x.contentStatus = xn.innerText

  if app.hasChild(NsExtendedProperties\"Application", xn): x.application = xn.innerText
  if app.hasChild(NsExtendedProperties\"Manager", xn): x.manager = xn.innerText
  if app.hasChild(NsExtendedProperties\"Company", xn): x.company = xn.innerText

proc save(x: XlProperties, core: XlNode, app: XlNode): tuple[core: XlNode, app: XlNode] =
  result.core = core.dup()
  result.app = app.dup()

  result.core.addNameSpace(NsCoreProperties, "cp")
  result.core.addNameSpace(NsDc, "dc")
  result.core.addNameSpace(NsDcTerms, "dcterms")
  result.core.addNameSpace(NsXSI, "xsi")
  result.app.addNameSpace(NsExtendedProperties)

  template setText(xn: XlNode, tag: XlNsName, text: string): untyped =
    if text != "":
      xn.replace(xn.newXlNode(tag, text))
    else:
      xn.delete(tag)

  result.core.setText(NsDc\"title", x.title)
  result.core.setText(NsDc\"subject", x.subject)
  result.core.setText(NsDc\"creator", x.creator)
  result.core.setText(NsCoreProperties\"keywords", x.keywords)
  result.core.setText(NsDc\"description", x.description)
  result.core.setText(NsCoreProperties\"lastModifiedBy", x.lastModifiedBy)
  result.core.setText(NsDcTerms\"created", x.created)
  result.core.setText(NsDcTerms\"modified", x.modified)
  result.core.setText(NsCoreProperties\"category", x.category)
  result.core.setText(NsCoreProperties\"contentStatus", x.contentStatus)

  result.app.setText(NsExtendedProperties\"Application", x.application)
  result.app.setText(NsExtendedProperties\"Manager", x.manager)
  result.app.setText(NsExtendedProperties\"Company", x.company)

  if result.core.hasChild(NsDcTerms\"created", xn): xn[NsXSI\"type"] = "dcterms:W3CDTF"
  if result.core.hasChild(NsDcTerms\"modified", xn): xn[NsXSI\"type"] = "dcterms:W3CDTF"

proc load(x: var XlSheetProtection, xProtection: XlNode) =
  template loadOption(x, sym: untyped): untyped =
    x.sym = xProtection.hasBoolOptionAttr(astToStr(sym))

  x.loadOption(sheet)
  x.loadOption(objects)
  x.loadOption(scenarios)
  x.loadOption(formatCells)
  x.loadOption(formatColumns)
  x.loadOption(formatRows)
  x.loadOption(insertColumns)
  x.loadOption(insertRows)
  x.loadOption(insertHyperlinks)
  x.loadOption(deleteColumns)
  x.loadOption(deleteRows)
  x.loadOption(selectLockedCells)
  x.loadOption(sort)
  x.loadOption(autoFilter)
  x.loadOption(pivotTables)
  x.loadOption(selectUnlockedCells)

  if xProtection.hasAttr("password", password) and password != "":
    x.password = some password

proc save(x: XlSheetProtection, xProtection: XlNode) =
  template saveOption(x, xn, sym: untyped): untyped =
    if x.sym.isSome:
      xProtection[astToStr(sym)] = if x.sym.get: "1" else: "0"
    else:
      xProtection.deleteAttr(astToStr(sym))

  x.saveOption(xProtection, sheet)
  x.saveOption(xProtection, objects)
  x.saveOption(xProtection, scenarios)
  x.saveOption(xProtection, formatCells)
  x.saveOption(xProtection, formatColumns)
  x.saveOption(xProtection, formatRows)
  x.saveOption(xProtection, insertColumns)
  x.saveOption(xProtection, insertRows)
  x.saveOption(xProtection, insertHyperlinks)
  x.saveOption(xProtection, deleteColumns)
  x.saveOption(xProtection, deleteRows)
  x.saveOption(xProtection, selectLockedCells)
  x.saveOption(xProtection, sort)
  x.saveOption(xProtection, autoFilter)
  x.saveOption(xProtection, pivotTables)
  x.saveOption(xProtection, selectUnlockedCells)

  # password
  #   none: keep old password
  #   some "": clear all the password
  #   some not "": set the new password
  if x.password.isSome:
    if x.password.get != "":
      xProtection["password"] = x.password.get
    else:
      xProtection.deleteAttr("password")
    xProtection.deleteAttr("algorithmName")
    xProtection.deleteAttr("hashValue")
    xProtection.deleteAttr("saltValue")
    xProtection.deleteAttr("spinCount")

proc getRootXlNode(xs: XlSheet): XlNode =
  var x: XmlParser
  open(x, newStringStream(xs.obj.raw), "xml", {reportWhitespace})
  x.next()
  defer: x.close()

  if not x.find(xmlElementOpen):
    raise newException(XlError, "XML parsing error")

  return x.parseIgnoreChild()

proc parseRow(xs: XlSheet, xn: XlNode): XlRow =
  if xn.hasIntAttr("r", r):
    r.dec()
    var xr = XlRow(sheet: xs, row: r, height: -1, style: -1)
    if xn.hasTrueAttr("customFormat") and xn.hasIntAttr("s", s): xr.style = s
    if xn.hasTrueAttr("customHeight") and xn.hasFloatAttr("ht", ht): xr.height = ht
    if xn.hasTrueAttr("hidden"): xr.hidden = true
    if xn.hasIntAttr("outlineLevel", ol): xr.outlineLevel = ol
    if xn.hasTrueAttr("collapsed"): xr.collapsed = true
    if xn.hasTrueAttr("thickTop"): xr.thickTop = true
    if xn.hasTrueAttr("thickBot"): xr.thickBot = true
    if xn.hasTrueAttr("ph"): xr.ph = true
    return xr

proc parseCell(xs: XlSheet, xn: XlNode,
    fTag: string, vTag: string, isTag: string, readonly: bool): XlCell =

  if xn.hasAttr("r", r):
    let rc = r.rc
    var cell = XlCell(sheet: xs, rc: rc, style: -1, readonly: readonly)

    if xn.hasRawChild(fTag, f): cell.formula = f.innerText
    if xn.hasRawChild(vTag, v): cell.value = v.innerText
    if xn.hasIntAttr("s", s): cell.style = s

    case xn.attr("t")
    of "s":
      cell.ct = ctSharedString
      try:
        let index = parseInt(cell.value)
        cell.riches = xs.workbook.sharedStrings.getRiches(index)
        cell.value = xs.workbook.sharedStrings.get(index)
        if cell.riches != default(XlRiches):
          cell.ct = ctInlineStr
          cell.value = cell.riches.value

      except ValueError:
        # cannot parse shared string index, let it be raw value
        discard

    of "b": cell.ct = ctBoolean
    of "e": cell.ct = ctError
    of "str": cell.ct = ctStr
    of "inlineStr":
      cell.ct = ctInlineStr
      if xn.hasRawChild(isTag, xIs):
        cell.riches.load(xIs)
        cell.value = cell.riches.value

    else: # n or other
      cell.ct = ctNumber

    return cell

proc extractSheetData(xs: XlSheet, root: XlNode, remove: bool): string =
  let
    sheetDataTag = nsnameToRawName(root, NsMain\"sheetData")
    tagOpen = '<' & sheetDataTag
    tagEnd = '/' & sheetDataTag & '>'
    first = xs.obj.raw.find(tagOpen)
    last = xs.obj.raw.find(tagEnd)

  if first != -1 and last != -1:
    result = xs.obj.raw[first..<(last + tagEnd.len)]
    if remove:
      xs.obj.raw = xs.obj.raw[0..<first] & xs.obj.raw[(last + tagEnd.len)..^1]

proc parseExperimental(xs: XlSheet) =
  assert xs.obj.raw != ""

  let
    root = xs.getRootXlNode()
    sheetData = xs.extractSheetData(root, remove=true)
    rowTag = nsnameToRawName(root, NsMain\"row")
    cTag = nsnameToRawName(root, NsMain\"c")
    fTag = nsnameToRawName(root, NsMain\"f")
    vTag = nsnameToRawName(root, NsMain\"v")
    isTag = nsnameToRawName(root, NsMain\"is")

  if sheetData == "": # <sheetData /> or no sheetData
    return

  var
    x: XmlParser
    errors: seq[string]

  open(x, newStringStream(sheetData), "xml", {reportWhitespace})
  x.next()
  defer: x.close()

  while true:
    let index = x.find([cTag, rowTag])
    if index < 0: break

    case index:
    of 1: # rowTag
      let xn = x.parseIgnoreChild()
      if xn != nil:
        let row = parseRow(xs, xn)
        if row != nil:
          xs.rows[row.row] = row

    of 0: # cTag
      var xn = x.parse(errors)
      if xn != nil:
        let cell = parseCell(xs, xn, fTag, vTag, isTag, readonly=false)
        if cell != nil:
          xs.update(cell)

    else: discard

proc parse(xs: XlSheet) =
  if xs.workbook.experimental:
    xs.parseExperimental()

  xs.obj.parse()
  let xSheet = xs.obj.xn

  if xSheet.hasChild(NsMain\"sheetData", sheetData):
    for xRow in sheetData.find(NsMain\"row"):
      if xRow.hasIntAttr(NsMain\"r", r):
        r.dec()
        var row = XlRow(sheet: xs, row: r, height: -1, style: -1)
        if xRow.hasTrueAttr(NsMain\"customFormat") and xRow.hasIntAttr(NsMain\"s", s): row.style = s
        if xRow.hasTrueAttr(NsMain\"customHeight") and xRow.hasFloatAttr(NsMain\"ht", ht): row.height = ht
        if xRow.hasTrueAttr(NsMain\"hidden"): row.hidden = true
        if xRow.hasIntAttr(NsMain\"outlineLevel", ol): row.outlineLevel = ol
        if xRow.hasTrueAttr(NsMain\"collapsed"): row.collapsed = true
        if xRow.hasTrueAttr(NsMain\"thickTop"): row.thickTop = true
        if xRow.hasTrueAttr(NsMain\"thickBot"): row.thickBot = true
        if xRow.hasTrueAttr(NsMain\"ph"): row.ph = true
        xs.rows[r] = row

      for c in xRow.find(NsMain\"c"):
        if c.hasAttr(NsMain\"r", r):
          let rc = r.rc
          var cell = XlCell(sheet: xs, rc: rc, style: -1)
          if c.hasChild(NsMain\"f", f): cell.formula = f.innerText
          if c.hasChild(NsMain\"v", v): cell.value = v.innerText
          if c.hasIntAttr(NsMain\"s", s): cell.style = s

          case c[NsMain\"t"]
          of "s":
            cell.ct = ctSharedString
            try:
              let index = parseInt(cell.value)
              cell.riches = xs.workbook.sharedStrings.getRiches(index)
              cell.value = xs.workbook.sharedStrings.get(index)
              if cell.riches != default(XlRiches):
                cell.ct = ctInlineStr
                cell.value = cell.riches.value

            except ValueError:
              # cannot parse shared string index, let it be raw value
              discard

          of "b": cell.ct = ctBoolean
          of "e": cell.ct = ctError
          of "str": cell.ct = ctStr
          of "inlineStr":
            cell.ct = ctInlineStr
            if c.hasChild(NsMain\"is", xIs):
              cell.riches.load(xIs)
              cell.value = cell.riches.value

          else: # n or other
            cell.ct = ctNumber

          xs.update(cell)

  if xSheet.hasChild(NsMain\"cols", xCols):
    for xCol in xCols.find(NsMain\"col"):
      if xCol.hasIntAttr(NsMain\"min", min) and xCol.hasIntAttr(NsMain\"max", max):
        var
          width = float -1
          style = -1
          hidden = false

        if xCol.hasIntAttr(NsMain\"style", s): style = s
        if xCol.hasTrueAttr(NsMain\"hidden"): hidden = true
        # customWidth has no effect?
        if xCol.hasFloatAttr(NsMain\"width", w): width = w
        for i in min..max:
          var index = i - 1
          var col = XlCol(sheet: xs, col: index,
            hidden: hidden, width: width, style: style)
          xs.cols[index] = col

  if xSheet.hasChild(NsMain\"mergeCells", xMergeCells):
    for xMergeCell in xMergeCells.find(NsMain\"mergeCell"):
      if xMergeCell.hasAttr(NsMain\"ref", r):
        xs.merges.incl r.rcrc

  if xSheet.hasChild(NsMain\"sheetPr", xSheetPr):
    if xSheetPr.hasChild(NsMain\"tabColor", xTabColor):
      xs.color = parseColor(xTabColor)

  if xSheet.hasChild(NsMain\"hyperlinks", xHyperlinks):
    for xHyperlink in xHyperlinks.find(NsMain\"hyperlink"):
      let
        rid = xHyperlink[NsDocumentRelationships\"id"]
        rc = xHyperlink[NsMain\"ref"].rc
        obj = xs.rels.getOrDefault(rid)

      if obj.external:
        xs.cells.mgetOrPut(rc,
          XlCell(sheet: xs, rc: rc, style: -1)
        ).hyperlink = obj.target

  if xSheet.hasChild(NsMain\"sheetProtection", xSheetProtection):
    xs.protection.load(xSheetProtection)

proc saveSheetDataExperimental(xw: XlWorkbook, xs: XlSheet, xSheetData: XlNode,
    rowCells: var Table[int, seq[XlCell]], hyperlinks: var seq[(string, string)]) =

  var sheetData = "<sheetData>"
  defer:
    sheetData.add "</sheetData>"
    xs.obj.postProcess = ("<sheetData />", sheetData)

  for r, cells in rowCells.sortedPairs:
    var rowChildren = ""
    sheetData.add "<row r=\"" & $(r + 1) & "\""
    defer:
      sheetData.add ">"
      sheetData.add rowChildren
      sheetData.add "</row>"

    let row = xs.rows.getOrDefault(r, nil)
    if not row.isNil:
      let style = xw.sharedStyles.get(row)
      if style >= 0:
        sheetData.add " customFormat=\"1\" s=\"" & $style & "\""

      if row.height >= 0:
        sheetData.add " customHeight=\"1\" ht=\"" & $row.height & "\""

      if row.outlineLevel != 0:
        sheetData.add " outlineLevel=\"" & $row.outlineLevel & "\""

      if row.hidden: sheetData.add " hidden=\"1\""
      if row.collapsed: sheetData.add " collapsed=\"1\""
      if row.thickTop: sheetData.add " thickTop=\"1\""
      if row.thickBot: sheetData.add " thickBot=\"1\""
      if row.ph: sheetData.add " ph=\"1\""

    cells.sort do (a, b: XlCell) -> int:
      system.cmp(a.rc, b.rc)

    # write cells data
    # should avoid write empty cell node
    for cell in cells:
      let name = cell.rc.name
      var
        modified = false
        xml = "<c r=\"" & name & "\""
        cellChildren = ""

      defer:
        if modified:
          rowChildren.add xml & '>'
          rowChildren.add cellChildren
          rowChildren.add "</c>"

      let style = xw.sharedStyles.get(cell)
      if style >= 0:
        xml.add " s=\"" & $style & "\""
        modified = true

      if cell.formula != "":
        cellChildren.add "<f>" & cell.formula & "</f>"
        modified = true

      if cell.ct == ctInlineStr:
        let xn = cell.riches.save(xSheetData, NsMain\"is")
        xml.add " t=\"" & $cell.ct & "\""
        cellChildren.add $$xn
        modified = true

      else:
        if cell.value != "":
          modified = true
          if cell.ct != ctNumber:
            xml.add " t=\"" & $cell.ct & "\""

          if cell.ct == ctSharedString:
            let index = xw.sharedStrings.add(cell.value)
            cellChildren.add "<v>" & $index & "</v>"

          else:
            cellChildren.add "<v>" & cell.value & "</v>"

      if cell.hyperlink != "":
        let id = xs.rels.add(TypeHyperlink, cell.hyperlink, true)
        hyperlinks.add (name, id)


proc saveSheetData(xw: XlWorkbook, xs: XlSheet, xSheetData: XlNode,
    rowCells: var Table[int, seq[XlCell]], hyperlinks: var seq[(string, string)]) =

  for r, cells in rowCells.sortedPairs:
    var xRow = xSheetData.addChildXlNode(NsMain\"row", -1, {"r": $(r + 1)})

    # write rows data
    let row = xs.rows.getOrDefault(r, nil)
    if not row.isNil:
      let style = xw.sharedStyles.get(row)
      if style >= 0:
        xRow["customFormat"] = "1"
        xRow["s"] = $style

      if row.height >= 0:
        xRow["customHeight"] = "1"
        xRow["ht"] = $row.height

      if row.outlineLevel != 0:
        xRow["outlineLevel"] = $row.outlineLevel

      if row.hidden: xRow["hidden"] = "1"
      if row.collapsed: xRow["collapsed"] = "1"
      if row.thickTop: xRow["thickTop"] = "1"
      if row.thickBot: xRow["thickBot"] = "1"
      if row.ph: xRow["ph"] = "1"

    cells.sort do (a, b: XlCell) -> int:
      system.cmp(a.rc, b.rc)

    # write cells data
    # should avoid write empty cell node
    for cell in cells:
      let name = cell.rc.name
      var xCell = xRow.newXlNode(NsMain\"c", {"r": name})
      var modified = false
      defer:
        if modified:
          xRow.add xCell

      let style = xw.sharedStyles.get(cell)
      if style >= 0:
        xCell["s"] = $style
        modified = true

      if cell.formula != "":
        discard xCell.addChildXlNode(NsMain\"f", cell.formula)
        modified = true

      if cell.ct == ctInlineStr:
        xCell["t"] = $cell.ct
        xCell.add cell.riches.save(xCell, NsMain\"is")
        modified = true

      else:
        if cell.value != "":
          modified = true
          if cell.ct != ctNumber:
            xCell["t"] = $cell.ct

          if cell.ct == ctSharedString:
            let index = xw.sharedStrings.add(cell.value)
            discard xCell.addChildXlNode(NsMain\"v", $index)

          else:
            discard xCell.addChildXlNode(NsMain\"v", cell.value)

      if cell.hyperlink != "":
        let id = xs.rels.add(TypeHyperlink, cell.hyperlink, true)
        hyperlinks.add (name, id)

proc saveSheets(xw: XlWorkbook) =
  const order = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr",
      "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
      "scenarios", "autoFilter", "sortState", "dataConsolidate",
      "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
      "dataValidations", "hyperlinks", "printOptions", "pageMargins",
      "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties",
      "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
      "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems",
      "tableParts", "extLst"]

  for index, xs in xw.sheets:
    # no parsing, no change
    if not xs.isParsed: continue

    # replace old sheet instead of create new one for save memory usage
    # children in sheet must be in order
    var xSheet = xs.obj.xn
    xSheet.addNameSpace(NsMain)
    xSheet.addNameSpace(NsDocumentRelationships, "r")
    xSheet.expand(NsMain, order)

    defer:
      xSheet.shrink(NsMain, ["sheetData"])

    # remove old hyperlinks
    xs.rels.removeExternal()
    var
      xSheetData = xSheet.newXlNode(NsMain\"sheetData")
      rowCells: Table[int, seq[XlCell]]
      hyperlinks: seq[(string, string)]

    xSheet.replace(xSheetData)

    # collect rows from both xs.cells and xs.rows
    for cell in xs.cells.values:
      rowCells.mgetOrPut(cell.rc.row, @[]).add cell

    for row in xs.rows.keys:
      discard rowCells.mgetOrPut(row, @[])

    if xw.experimental:
      saveSheetDataExperimental(xw, xs, xSheetData, rowCells, hyperlinks)
    else:
      saveSheetData(xw, xs, xSheetData, rowCells, hyperlinks)

    # write cols data
    type ColData = tuple[style: int, width: float, hidden: bool]
    if xs.cols.len != 0:
      var xCols = xSheet.newXlNode(NsMain\"cols")
      xSheet.replace(xCols)

      var last = ColData (-1, -1.0, false)
      for col in xs.cols.sortedValues:
        var
          n = $(col.col + 1)
          xCol = xCols.newXlNode(NsMain\"col", {"min": n, "max": n})
          modified = false
          current = ColData (-1, -1.0, false)

        defer:
          if modified:
            if current == last and xCols.count != 0 and xCols[^1]["max"] == $(col.col):
              xCols[^1]["max"] = n
            else:
              xCols.add(xCol)

          last = current

        let style = xw.sharedStyles.get(col)
        if style >= 0:
          current.style = style
          xCol["style"] = $style
          modified = true

        if col.width >= 0:
          current.width = col.width
          xCol["width"] = $col.width
          # customWidth has no effect?
          # xCol["customWidth"] = "1"
          modified = true
        else:
          xCol["width"] = "8.43"
          modified = true

        if col.hidden:
          current.hidden = col.hidden
          xCol["hidden"] = "1"
          modified = true

    else:
      xSheet.delete(NsMain\"cols")

    # write merges data
    if xs.merges.len != 0:
      var xMergeCells = xSheet.newXlNode(NsMain\"mergeCells")
      for rcrc in xs.merges:
        discard xMergeCells.addChildXlNode(NsMain\"mergeCell", -1, {"ref": rcrc.name})

      xMergeCells.updateCount()
      xSheet.replace(xMergeCells)

    else:
      xSheet.delete(NsMain\"mergeCells")

    # write hyperlinks data
    if hyperlinks.len != 0:
      var xHyperlinks = xSheet.newXlNode(NsMain\"hyperlinks")
      for (name, id) in hyperlinks:
        discard xHyperlinks.addChildXlNode(NsMain\"hyperlink", -1,
          {"ref": name, "r:id": id}
        )

      xSheet.replace(xHyperlinks)

    else:
      xSheet.delete(NsMain\"hyperlinks")

    # write sheetProtection
    xs.protection.save(xSheet.child(NsMain\"sheetProtection"))

    # wreite tabColor in sheetPr
    if xs.color != default(XlColor):
      validate(xs.color, error=true)
      xSheet.editChild(NsMain\"sheetPr", xSheetPr, 0):
        xSheetPr.replace(xSheetPr.newChildColor(NsMain\"tabColor", xs.color))

    # clear tabSelected
    if index != xw.active and xSheet.hasChild(NsMain\"sheetViews", xSheetViews):
      for xSheetView in xSheetViews.find(NsMain\"sheetView"):
        xSheetView.deleteAttr("tabSelected")

proc saveSheetRecords(xw: XlWorkbook, mime = MimeSheetMain): seq[(string, XlObject)] =
  # write sheet record into Workbook, WorkbookRels, and ContentTypes

  const order = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection",
    "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames",
    "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
    "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects",
    "extLst" ]

  var xWorkbook = xw{Workbook}.dup()
  xWorkbook.addNameSpace(NsMain)
  xWorkbook.addNameSpace(NsDocumentRelationships, "r")
  xWorkbook.expand(NsMain, order)

  var xSheets = xWorkbook.newXlNode(NsMain\"sheets")
  xWorkbook.replace(xSheets)

  var
    rels: XlRels
    types: XlTypes

  for i, sheet in xw.sheets:
    let
      id = i + 1
      target = "worksheets/sheet" & $id & ".xml"

    result.add (Xl/target, sheet.obj)
    if sheet.rels.len != 0:
      result.add (Xl/target.getRelsFile, newXlObject(sheet.rels.save()))

    let rid = rels.add(TypeWorksheet, Xl/target)
    types.add(/Xl/target, MimeWorksheet, true)

    var xSheet = xSheets.addChildXlNode(NsMain\"sheet", -1,
      {"name": sheet.name, "sheetId": $id, "r:id": rid}
    )
    if sheet.hidden:
      xSheet["state"] = "hidden"

  for rel in xw.rels.sortedValues:
    if rel.typ != TypeWorksheet:
      # the rid is not used if not a sheet
      rels.add rel

  for key, obj in xw.types.sortedPairs:
    if obj.typ == MimeCalcChain:
      # remove calcChain to avoid open failure
      # todo: fixme
      continue

    if obj.typ != MimeWorksheet:
      types.add(key, obj)

  result.add (Workbook.getRelsFile, newXlObject(rels.save()))
  result.add (ContentTypes, newXlObject(types.save(mime)))

  xWorkbook.editChild(NsMain\"bookViews", xBookViews, -1):
    xBookViews.editChild(NsMain\"workbookView", xWorkbookView, -1):
      xWorkbookView["activeTab"] = $xw.active

  # force to recalculate formula
  xWorkbook.editChild(NsMain\"calcPr", xCalcPr, -1):
    xCalcPr["fullCalcOnLoad"] = "1"
    if xCalcPr["calcId"] == "":
      xCalcPr["calcId"] = "124519"

  xWorkbook.shrink(NsMain, ["sheets"])
  result.add (Workbook, newXlObject(xWorkbook))

proc saveContents(xw: XlWorkbook, mime = MimeSheetMain): Table[string, XlObject] =
  result = xw.contents

  xw.saveSheets()
  result[SharedStrings] = newXlObject(xw.sharedStrings.save())
  result[Styles] = newXlObject(xw.sharedStyles.save())

  xw.sharedStrings.reset()
  xw.sharedStyles.reset()

  for (file, obj) in xw.saveSheetRecords(mime):
    result[file] = obj

  let (core, app) = xw.properties.save(xw{DocPropsCore}, xw{DocPropsApp})
  result[DocPropsCore] = newXlObject(core)
  result[DocPropsApp] = newXlObject(app)

proc save*(xw: XlWorkbook, path: string) =
  ## Save XlWorkbook object into a xlsx or xltx file.
  if xw.sheets.len == 0:
    raise newException(XlError, "No sheet in this workbook.")

  let mime = if path.splitFile.ext.toLowerAscii == ".xltx": MimeTemplateMain
    else: MimeSheetMain

  let
    contents = xw.saveContents(mime)
    archive = ZipArchive()

  for file in contents.sortedKeys:
    archive.contents[file] = ArchiveEntry(kind: ekFile, contents: $contents[file])

  archive.writeZipArchive(path)

proc createStyles(xw: XlWorkbook) =
  xw.rels.add(TypeStyles, "styles.xml")
  xw.types.add(/Styles, MimeStyles, true)

  var xStyleSheet = newXlRootNode(NsMain\"styleSheet")
  var xFonts = xStyleSheet.addChildXlNode(NsMain\"fonts")
  discard xFonts.addChildXlNode(NsMain\"font")
  xFonts.updateCount()

  var xFills = xStyleSheet.addChildXlNode(NsMain\"fills")
  var xFill: XlNode
  xFill = xFills.addChildXlNode(NsMain\"fill")
  discard xFill.addChildXlNode(NsMain\"patternFill", -1, {"patternType": "none"})
  xFill = xFills.addChildXlNode(NsMain\"fill")
  discard xFill.addChildXlNode(NsMain\"patternFill", -1, {"patternType": "none"})
  xFills.updateCount()

  var xBorders = xStyleSheet.addChildXlNode(NsMain\"borders")
  discard xBorders.addChildXlNode(NsMain\"border")
  xFills.updateCount()

  var xCellXfs = xStyleSheet.addChildXlNode(NsMain\"cellXfs")
  discard xCellXfs.addChildXlNode(NsMain\"xf", -1,
    {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"}
  )
  xCellXfs.updateCount()

  xw.contents[Styles] = newXlObject(xStyleSheet)

proc createSharedStrings(xw: XlWorkbook) =
  xw.rels.add(TypeSharedStrings, "sharedStrings.xml")
  xw.types.add(/SharedStrings, MimeSharedStrings, true)

  xw.contents[SharedStrings] = newXlObject(newXlRootNode(NsMain\"sst"))

proc createProperties(xw: XlWorkbook) =
  if DocPropsCore notin xw.contents:
    var xCoreProp = newXlRootNode(NsCoreProperties\"coreProperties", "cp")
    xw.contents[DocPropsCore] = newXlObject(xCoreProp)

  if DocPropsApp notin xw.contents:
    var xExtendedProp = newXlRootNode(NsExtendedProperties\"Properties")
    xw.contents[DocPropsApp] = newXlObject(xExtendedProp)

proc load*(path: string, experimental = false): XlWorkbook =
  ## Load xlsx or xltx file as XlWorkbook object.
  result = XlWorkbook(experimental: experimental)
  let reader = openZipArchive(path)
  for file in reader.walkFiles:
    result.contents[file] = newXlObject(reader.extractFile(file))
  reader.close()

  if ContentTypes in result.contents:
    result.types.load(result{ContentTypes})

  let workbookRels = Workbook.getRelsFile
  if workbookRels in result.contents:
    result.rels.load(result{workbookRels}, workbookRels)

  var xWorkbook: XlNode
  try:
    xWorkbook = result{Workbook}
  except:
    raise newException(XlError, "Invaild Open XML file")

  if xWorkbook.hasChild(NsMain\"sheets", xSheets):
    for xSheet in xSheets.find(NsMain\"sheet"):
      # r:id in normal
      let id = xSheet[NsDocumentRelationships\"id"]
      if id notin result.rels:
        continue

      var target = result.rels[id].target
      if target notin result.contents:
        continue

      var sheet = XlSheet(
        workbook: result,
        name: xSheet["name"],
        hidden: xSheet["state"] == "hidden",
        obj: result.contents[target],
        first: EmptyRowCol,
        last: EmptyRowCol
      )
      result.sheets.add sheet
      result.contents.del target

      let targetRel = target.getRelsFile()
      if targetRel in result.contents:
        sheet.rels.load(result{targetRel}, targetRel)
        result.contents.del targetRel

  if xWorkbook.hasChild(NsMain\"bookViews", xBookViews):
    if xBookViews.hasChild(NsMain\"workbookView", xWorkbookView):
      try:
        result.active = parseInt(xWorkbookView["activeTab"])
        if result.active >= result.sheets.len:
          result.active = 0
      except: discard

  if SharedStrings notin result.contents: result.createSharedStrings()
  if Styles notin result.contents: result.createStyles()
  result.sharedStrings.load(result{SharedStrings})
  result.sharedStyles.load(result{Styles})

  result.createProperties() # ensure DocPropsCore and DocPropsApp exists
  result.properties.load(result{DocPropsCore}, result{DocPropsApp})

proc standardizeIndex(xw: XlWorkbook, index: int): int {.inline.} =
  result = if index < 0: xw.sheets.len + index else: index

  if result notin 0..xw.sheets.high:
    raise newException(XlError, "index out of range")

proc contains*(rcrc: XlRCRC, rc: XlRC): bool {.inline.} =
  ## Check if a cell is in a range.
  result = rc.row in rcrc[0].row..rcrc[1].row and
    rc.col in rcrc[0].col..rcrc[1].col

proc contains*(xr: XlRange, rc: XlRC|string): bool {.inline.} =
  ## Check if a cell is in a range.
  when rc is string:
    let rc = rc.rc

  elif not compiles(rc.row):
    let rc = XlRC rc

  result = rc.row in xr.first.row..xr.last.row and
    rc.col in xr.first.col..xr.last.col

proc contains*(xco: XlCollection, rc: XlRC|string): bool {.inline.} =
  ## Check if a cell is in a collection.
  when rc is string:
    let rc = rc.rc

  result = rc in xco.rcs

proc contains*(xr: XlRange, xc: XlCell): bool {.inline.} =
  ## Check if a cell is in a range.
  result = xc.sheet == xr.sheet and xr.contains(xc.rc)

proc contains*(xco: XlCollection, xc: XlCell): bool {.inline.} =
  ## Check if a cell is in a collection.
  result = xc.sheet == xco.sheet and xco.contains(xc.rc)

proc contains*(xs: XlSheet, xc: XlCell): bool {.inline.} =
  ## Check if a cell is in a sheet.
  result = xc.sheet == xs and xc.rc in xs.cells

proc contains*(xs: XlSheet, rc: XlRC): bool {.inline.} =
  ## Check if a cell is in a sheet.
  result = rc in xs.cells

proc newSheet(xw: XlWorkbook, name: string): XlSheet =
  var xSheet = newXlRootNode(NsMain\"worksheet")
  xSheet.addNameSpace(NsDocumentRelationships, "r")
  discard xSheet.addChildXlNode(NsMain\"sheetData")

  result = XlSheet(
    workbook: xw,
    name: name,
    obj: newXlObject(xSheet),
    first: EmptyRowCol,
    last: EmptyRowCol
  )

proc newWorkbook*(experimental = false): XlWorkbook =
  ## Create a new workbook.
  result = XlWorkbook(experimental: experimental)

  result.types.add("rels", MimeRelationships, false)
  result.types.add("xml", MimeXml, false)
  result.types.add(/DocPropsCore, MimeCoreProperties, true)
  result.types.add(/DocPropsApp, MimeExtendedProperties, true)
  result.types.add(/Workbook, MimeSheetMain, true)

  var rels: XlRels
  rels.add(TypeOfficeDocument, Workbook)
  rels.add(TypeCoreProperties, DocPropsCore)
  rels.add(TypeExtendedProperties, DocPropsApp)
  result.contents["_rels/.rels"] = newXlObject(rels.save())

  var xWorkbook = newXlRootNode(NsMain\"workbook")
  xWorkbook.addNameSpace(NsDocumentRelationships, "r")
  discard xWorkbook.addChildXlNode(NsMain\"sheets")
  result.contents[Workbook] = newXlObject(xWorkbook)

  result.createSharedStrings()
  result.createStyles()
  result.createProperties()
  result.sharedStrings.load(result{SharedStrings})
  result.sharedStyles.load(result{Styles})

proc add*(xw: XlWorkbook, name: string): XlSheet {.discardable.} =
  ## Add a new sheet into workbook.
  for s in xw.sheets:
    if name == s.name:
      raise newException(XlError, "Duplicated sheet name")

  let sheet = xw.newSheet(name)
  xw.sheets.add sheet
  return sheet

proc insert*(xw: XlWorkbook, name: string, index = 0): XlSheet {.discardable.} =
  ## Insert a new sheet into workbook at specified index.
  let index =
    if index == xw.sheets.len: index
    else: xw.standardizeIndex(index)

  for s in xw.sheets:
    if name == s.name:
      raise newException(XlError, "Duplicated sheet name")

  let sheet = xw.newSheet(name)
  xw.sheets.insert(sheet, index)
  return sheet

proc sheet*(xw: XlWorkbook, index: int): XlSheet =
  ## Return the sheet at specified index in a workbook.
  let index = xw.standardizeIndex(index)
  var sheet = xw.sheets[index]
  if not sheet.isParsed:
    sheet.parse()

  return sheet

proc sheet*(xw: XlWorkbook, name: string): XlSheet =
  ## Return the sheet of specified name in a workbook.
  for i, s in xw.sheets:
    if s.name == name:
      return xw.sheet(i)

  raise newException(XlError, name & " not found")

proc sheet*(x: XlCell|XlRow|XlCol|XlRange|XlCollection): XlSheet {.inline.} =
  ## Return the sheet of these objects.
  return x.sheet

proc `[]`*(xw: XlWorkbook, index: int): XlSheet {.inline.} =
  ## Return the sheet at specified index in a workbook.
  ## Syntactic sugar of sheet().
  return xw.sheet(index)

proc `[]`*(xw: XlWorkbook, name: string): XlSheet {.inline.} =
  ## Return the sheet of specified name in a workbook.
  ## Syntactic sugar of sheet().
  return xw.sheet(name)

proc active*(xw: XlWorkbook): XlSheet {.inline.} =
  ## Return the active sheet in a workbook.
  return xw.sheet(xw.active)

proc `active=`*(xw: XlWorkbook, index: int) =
  ## Set the sheet at specified index as active sheet.
  if index >= xw.sheets.len:
    raise newException(XlError, "index out of range")

  if index != xw.active:
    # force to parse the sheet so that
    # we can deal with tabSelected in sheetViews
    # is there a better way?
    discard xw.sheet(xw.active)

  xw.active = index

proc `active=`*(xw: XlWorkbook, name: string) =
  ## Set the sheet of specified name as active sheet.
  for i, s in xw.sheets:
    if s.name == name:
      xw.active = i
      return

  raise newException(XlError, name & " not found")

proc properties*(xw: XlWorkbook): var XlProperties {.inline.} =
  ## Return modifiable XlProperties var object of a workbook.
  return xw.properties

proc `properties=`*(xw: XlWorkbook, properties: XlProperties) {.inline.} =
  ## Set XlProperties object of a workbook.
  ## Assign default(XlProperties) to reset to default.
  xw.properties = properties

proc len*(xw: XlWorkbook): int {.inline.} =
  ## Return how many sheets in a workbook.
  result = xw.sheets.len

proc count*(xw: XlWorkbook): int {.inline.} =
  ## Return how many sheets in a workbook.
  result = xw.sheets.len

proc delete*(xw: XlWorkbook, index: int) {.inline.} =
  ## Delete a sheet at specified index in a workbook.
  let index = xw.standardizeIndex(index)
  xw.sheets.delete(index)

proc delete*(xw: XlWorkbook, name: string) =
  ## Delete a sheet of specified name in a workbook.
  for i, s in xw.sheets:
    if s.name == name:
      xw.sheets.delete(i)
      break

iterator sheetNames*(xw: XlWorkbook): string =
  ## Iterate over sheet names in a workbook.
  for s in xw.sheets:
    yield s.name

iterator cells*(xw: XlWorkbook, index: int): XlCell =
  ## The fastest way to iterate over cells of a sheet.
  ## The yield cells are readonly and empty cells may be skipped.
  ## Check `cell.rc` carefully during iteration to avoid problem.
  ## This may be changed in the future.
  let index = xw.standardizeIndex(index)
  var xs = xw.sheets[index]
  if xs.isParsed:
    raise newException(XlError, "Using cells iteraotr in parsed sheet is meaningless.")

  let
    root = xs.getRootXlNode()
    sheetData = xs.extractSheetData(root, remove=false)
    cTag = nsnameToRawName(root, NsMain\"c")
    fTag = nsnameToRawName(root, NsMain\"f")
    vTag = nsnameToRawName(root, NsMain\"v")
    isTag = nsnameToRawName(root, NsMain\"is")

  if sheetData != "":
    var
      x: XmlParser
      errors: seq[string]

    open(x, newStringStream(sheetData), "xml", {reportWhitespace})
    x.next()
    defer: x.close()

    while x.find([cTag]) >= 0:
      let xn = x.parse(errors)
      if xn != nil:
        let cell = parseCell(xs, xn, fTag, vTag, isTag, readonly=true)
        if cell != nil:
          yield cell

iterator cells*(xw: XlWorkbook, name: string): XlCell =
  ## The fastest way to iterate over cells of a sheet.
  ## The yield cells are readonly and empty cells may be skipped.
  ## Check `cell.rc` carefully during iteration to avoid problem.
  ## This may be changed in the future.
  var found = -1
  for i, s in xw.sheets:
    if s.name == name:
      found = i
      break

  if found >= 0:
    for cell in xw.cells(found):
      yield cell

  else:
    raise newException(XlError, name & " not found")

proc index*(xs: XlSheet): int {.inline.} =
  ## Return index of the sheet.
  result = xs.workbook.sheets.find xs
  assert result >= 0

proc swap*(xs: XlSheet, index: int) =
  ## Swap sheets.
  let
    xw = xs.workbook
    index = xw.standardizeIndex(index)
    oldIndex = xs.index()

  swap(xw.sheets[oldIndex], xw.sheets[index])

proc move*(xs: XlSheet, index: int) =
  ## Move sheet.
  let
    xw = xs.workbook
    index = xw.standardizeIndex(index)
    oldIndex = xs.index()

  if oldIndex == index:
    return

  let sheet = move(xw.sheets[oldIndex])
  if oldIndex > index:
    for i in countdown(oldIndex, index + 1):
      xw.sheets[i] = move(xw.sheets[i - 1])
  else:
    for i in oldIndex..<index:
      xw.sheets[i] = move(xw.sheets[i + 1])
  xw.sheets[index] = sheet

proc copy*(xs: XlSheet, name = ""): XlSheet {.discardable.} =
  ## Copy a sheet into new name and return it.
  ## Copy drawing of a sheet (include images and comments) is not support now.
  let
    xw = xs.workbook
    index = xs.index()

  var sheet = XlSheet(
    workbook: xw,
    name: if name == "": xs.name & "_copy" else: name,
    obj: XlObject(),
    first: EmptyRowCol,
    last: EmptyRowCol
  )

  if xs.obj.xn.isNil:
    sheet.obj.raw = xs.obj.raw

  else:
    sheet.obj.raw = $xs.obj.xn

  sheet.parse()

  # todo: copy drawing of a sheet?
  sheet.obj.xn.delete(NsMain\"drawing")
  sheet.obj.xn.delete(NsMain\"legacyDrawing")

  xw.sheets.insert(sheet, index + 1)
  return sheet

proc delete*(xs: XlSheet) {.inline.} =
  ## Delete a sheet.
  xs.workbook.sheets.delete(xs.index)

proc delete*(xs: XlSheet, c: XlRC|XlCell|string) {.inline.} =
  ## Delete a cell from a sheet.
  ## The old XlCell object will become useless, should not edit it anymore.
  when c is XlRC:
    let rc = c
  elif c is XlCell:
    let rc = c.rc
  else:
    let rc = c.rc
  xs.cells.del rc
  xs.update()

proc name*(xs: XlSheet): string {.inline.} =
  ## Return name of a sheet.
  result = xs.name

proc `name=`*(xs: XlSheet, name: string) =
  ## Rename a sheet.
  for s in xs.workbook.sheets:
    if s == xs: continue
    if s.name == name:
      raise newException(XlError, "Duplicated sheet name")

  xs.name = name

proc color*(xs: XlSheet): XlColor {.inline.} =
  ## Return tab color of a sheet.
  result = xs.color

proc `color=`*(xs: XlSheet, color: XlColor) {.inline.} =
  ## Set tab color of a sheet.
  ## Assign default(XlColor) to reset to default.
  xs.color = color

proc protection*(xs: XlSheet): var XlSheetProtection {.inline.} =
  ## Return modifiable XlSheetProtection var object of a sheet.
  result = xs.protection

proc `protection=`*(xs: XlSheet, protection: XlSheetProtection) {.inline.} =
  ## Set XlSheetProtection object of a sheet.
  ## Assign default(XlSheetProtection) to reset to default.
  xs.protection = protection

proc protect*(xs: XlSheet, enable = true, password = none(string)) =
  ## Protect a sheet by default setting of XlSheetProtection.
  ## By default, old password will be kept.
  ## If password is empyt string, old password will be removed.

  # password
  #   none: keep old password
  #   some "": clear all the password
  #   some not "": set the new password

  proc hash(password: string): uint16 {.used.} =
    var hash: uint32
    for i in 0..<password.len:
      var
        letter = password[i].ord.uint32 shl (i + 1)
        low = letter and 0x7fff
        high = (letter and (0x7fff shl 15)) shr 15

      hash = hash xor (low or high)

    return cast[uint16](hash) xor cast[uint16](password.len) xor 0xce4b

  if enable:
    xs.protection = XlSheetProtection(sheet: true, objects: true, scenarios: true)
    if password.isSome and password.get != "":
      xs.protection.password = some toHex(hash(password.get))
    else:
      xs.protection.password = password

  else:
    xs.protection = XlSheetProtection()

proc first*(xs: XlSheet|XlRange): XlRC {.inline.} =
  ## Return first cell of a sheet or range.
  result = xs.first

proc last*(xs: XlSheet|XlRange): XlRC {.inline.} =
  ## Return last cell of a sheet or range.
  result = xs.last

proc dimension*(xs: XlSheet|XlRange): XlRCRC {.inline.} =
  ## Return dimension of a sheet or range.
  if not xs.isEmpty:
    return (xs.first, xs.last)

proc cell*(xs: XlSheet, rc: XlRC): XlCell =
  ## Return a cell of a sheet.
  ## Indexed by XlRC tuple.
  if rc.row < 0 or rc.col < 0:
    raise newException(XlError, "invalid XlRowCol")

  xs.update(rc)

  # apply the style change of row/col setting
  proc applyChange(change: var XlChange, x: XlStylable) =
    template apply(symbol: untyped): untyped =
      if x.change.symbol.isSome and change.symbol.isNone:
        change.symbol = x.change.symbol

    if x != nil and x.change != nil:
      if change.isNil: change = XlChange()
      apply(numFmt)
      apply(font)
      apply(fill)
      apply(border)
      apply(alignment)
      apply(protection)

  var change: XlChange
  change.applyChange(xs.rows.getOrDefault(rc.row, nil))
  change.applyChange(xs.cols.getOrDefault(rc.col, nil))

  return xs.cells.mgetOrPut(rc, XlCell(sheet: xs, rc: rc, style: -1, change: change))

proc row*(xs: XlSheet, row: int): XlRow =
  ## Return a raw (as XlRow object) of a sheet.
  ## Index by int start from 0.
  if row < 0:
    raise newException(XlError, "invalid row")

  return xs.rows.mgetOrPut(row, XlRow(sheet: xs, row: row, height: -1, style: -1))

proc row*(xs: XlSheet, row: string): XlRow =
  ## Return a raw (as XlRow object) of a sheet.
  ## Index by cell reference (e.g. "A1" or "1")
  let rc = rc("A" & row)
  return xs.row(rc.row)

proc col*(xs: XlSheet, col: int): XlCol =
  ## Return a column (as XlCol object) of a sheet.
  ## Index by int start from 0.
  if col < 0:
    raise newException(XlError, "invalid col")

  return xs.cols.mgetOrPut(col, XlCol(sheet: xs, col: col, width: -1, style: -1))

proc col*(xs: XlSheet, col: string): XlCol =
  ## Return a column (as XlCol object) of a sheet.
  ## Index by cell reference (e.g. "A1" or "A")
  let rc = rc(col & "1")
  return xs.col(rc.col)

proc cell*(xs: XlSheet, name: string): XlCell {.inline.} =
  ## Return a cell of a sheet.
  ## Indexed by cell reference (e.g. "A1").
  result = xs.cell(name.rc)

proc cell*(xs: XlSheet, row, col: int): XlCell {.inline.} =
  ## Return a cell of a sheet.
  ## Indexed by row and col value start from 0.
  result = xs.cell((row, col))

proc cell*(row: XlRow, col: int): XlCell {.inline.} =
  ## Return a cell in a row of a sheet.
  ## Index by int start from 0.
  result = row.sheet.cell(row.row, col)

proc cell*(row: XlRow, col: string): XlCell {.inline.} =
  ## Return a cell in a row of a sheet.
  ## Index by cell reference (e.g. "A1" or "A")
  let rc = rc(col & "1")
  result = row.sheet.cell(row.row, rc.col)

proc cell*(col: XlCol, row: int): XlCell {.inline.} =
  ## Return a cell in a column of a sheet.
  ## Index by int start from 0.
  result = col.sheet.cell(row, col.col)

proc cell*(col: XlCol, row: string): XlCell {.inline.} =
  ## Return a cell in a column of a sheet.
  ## Index by cell reference (e.g. "A1" or "1")
  let rc = rc("A" & row)
  result = col.sheet.cell(rc.row, col.col)

proc range*(xs: XlSheet): XlRange {.inline.} =
  ## Return a range of a sheet.
  ## The returned range will include all the cells in the sheet.
  result = XlRange(sheet: xs, first: xs.first, last: xs.last)

proc range*(xs: XlSheet, first: XlRC|string, last: XlRC|string): XlRange =
  ## Return a arbitrary range of a sheet.
  ## Indexed by XlRC tuple or cell reference (e.g. "A1").
  when first is string:
    let first = first.rc

  when last is string:
    let last = last.rc

  if first < last:
    result = XlRange(sheet: xs, first: first, last: last)

  elif first >= last:
    result = XlRange(sheet: xs, first: last, last: first)

  if result.isEmpty:
    raise newException(XlError, "Cannot create an empyt range.")

proc range*(xs: XlSheet, rcrc: XlRCRC): XlRange {.inline.} =
  ## Return a arbitrary range of a sheet.
  ## Indexed by XlRCRC tuple.
  result = xs.range(rcrc.first, rcrc.last)

proc range*(xs: XlSheet, r: string): XlRange {.inline.} =
  ## Return a arbitrary range of a sheet.
  ## Indexed by range reference (e.g. "A1:C3").
  result = xs.range(r.rcrc)

proc range*(xr: XlRange, rcrc: XlRCRC): XlRange {.inline.} =
  ## Return a subrange range of a sheet.
  ## Indexed by XlRCRC tuple.

proc range*(xco: XlCollection): XlRange =
  ## Return a range include all the cells in a collection.
  result = XlRange(sheet: xco.sheet)
  if xco.rcs.len == 0:
    result.first = EmptyRowCol
    result.last = EmptyRowCol
  else:
    result.first = (int.high, int.high)
    result.last = (int.low, int.low)
    for rc in xco.rcs:
      result.updateFirstLast(rc)

proc row*(xr: XlRange, row: int): XlRange =
  ## Return a row (as XlRange object) in another range.
  ## Index by int start from 0.
  if row notin xr.first.row..xr.last.row:
    raise newException(XlError, "index not in " & $(xr.first.row..xr.last.row))

  result = XlRange(sheet: xr.sheet,
    first: (row, xr.first.col), last: (row, xr.last.col))

proc row*(xr: XlRange, row: string): XlRange =
  ## Return a row (as XlRange object) in another range.
  ## Index by cell reference (e.g. "A1" or "1")
  let rc = rc("A" & row)
  return xr.row(rc.row)

proc col*(xr: XlRange, col: int): XlRange =
  ## Return a column (as XlRange object) in another range.
  ## Index by int start from 0.
  if col > xr.last.col or col < xr.first.col:
    raise newException(XlError, "index not in " & $(xr.first.col..xr.last.col))

  result = XlRange(sheet: xr.sheet,
    first: (xr.first.row, col), last: (xr.last.row, col))

proc col*(xr: XlRange, col: string): XlRange =
  ## Return a column (as XlRange object) in another range.
  ## Index by cell reference (e.g. "A1" or "A")
  let rc = rc(col & "1")
  return xr.col(rc.col)

proc rowCount*(xr: XlRange): int {.inline.} =
  ## Return row counts of a range.
  if xr.isEmpty: return 0
  result = xr.last.row - xr.first.row + 1

proc colCount*(xr: XlRange): int {.inline.} =
  ## Return column counts of a range.
  if xr.isEmpty: return 0
  result = xr.last.col - xr.first.col + 1

proc count*(xr: XlRange): int {.inline.} =
  ## Return count of cells in a range.
  result = xr.rowCount * xr.colCount

proc len*(xr: XlRange): int {.inline.} =
  ## Return count of cells in a range.
  result = xr.count

proc cell*(xr: XlRange, index: int): XlCell =
  ## Return a cell in a range.
  ## Index by int start from 0.
  if xr.isEmpty:
    raise newException(XlError, "range is empty.")

  if index >= xr.count:
    raise newException(XlError, "index out of range")

  var rc = xr.first
  if xr.turned:
    rc.col += index div xr.rowCount
    rc.row += index mod xr.colCount
  else:
    rc.row += index div xr.colCount
    rc.col += index mod xr.rowCount
  return xr.sheet.cell(rc)

proc cell*(xr: XlRange, rc: XlRC): XlCell {.inline.} =
  ## Return a cell in a range (row, col relative to first cell).
  return xr.sheet.cell(xr.first.row + rc.row, xr.first.col + rc.col)

proc cell*(xr: XlRange, row: int, col: int): XlCell {.inline.} =
  ## Return a cell in a range (row, col relative to first cell).
  return xr.sheet.cell(xr.first.row + row, xr.first.col + col)

proc `[]`*(xs: XlSheet, rc: XlRC|string): XlCell {.inline.} =
  ## Return a cell of a sheet.
  ## Indexed by XlRC tuple or cell reference (e.g. "A1").
  ## Syntactic sugar of cell().
  result = xs.cell(rc)

proc `[]`*(xs: XlSheet, row, col: int): XlCell {.inline.} =
  ## Return a cell of a sheet.
  ## Indexed by row and col.
  ## Syntactic sugar of cell().
  result = xs.cell(row, col)

proc `[]`*(x: XlRow|XlCol, index: int|string): XlCell {.inline.} =
  ## Return a cell of a sheet.
  ## Syntactic sugar of cell().
  result = x.cell(index)

proc `[]`*(xs: XlSheet, rcrc: XlRCRC): XlRange {.inline.} =
  ## Return a arbitrary range of a sheet.
  ## Indexed by XlRCRC tuple.
  ## Syntactic sugar of range().
  result = xs.range(rcrc)

proc `{}`*(xs: XlSheet, rcrc: XlRCRC|string): XlRange {.inline.} =
  ## Return a arbitrary range of a sheet.
  ## Indexed by XlRCRC tuple or range reference (e.g. "A1:C3")
  ## Syntactic sugar of range().
  result = xs.range(rcrc)

proc `{}`*(xs: XlSheet, first: XlRC|string, last: XlRC|string): XlRange {.inline.} =
  ## Return a arbitrary range of a sheet.
  ## Syntactic sugar of range().
  result = xs.range(first, last)

proc `[]`*(x: XlRow, col: int|string): XlCell {.inline.} =
  ## Return a cell in a row of a sheet.
  ## Index by int start from 0 or cell reference (e.g. "A1" or "A").
  ## Syntactic sugar of cell().
  return x.cell(col)

proc `[]`*(x: XlCol, row: int|string): XlCell {.inline.} =
  ## Return a cell in a column of a sheet.
  ## Index by int start from 0 or cell reference (e.g. "A1" or "1").
  ## Syntactic sugar of cell().
  return x.cell(row)

proc `[]`*(xr: XlRange, index: int|XlRC): XlCell {.inline.} =
  ## Return a cell in a range.
  ## Syntactic sugar of cell().
  return xr.cell(index)

proc `[]`*(xr: XlRange, row: int, col: int): XlCell {.inline.} =
  ## Return a cell in a range.
  ## Syntactic sugar of cell().
  return xr.cell(row, col)

iterator merges*(xs: XlSheet): XlRange =
  ## Iterate over merged ranges in a sheet.
  for rcrc in xs.merges.sortedItems:
    yield XlRange(sheet: xs, first: rcrc[0], last: rcrc[1])

iterator rows*(xs: XlSheet, restrict = false): XlRow =
  ## Iterate over rows in a sheet. Yield XlRow objects.
  ## Both rows in dimension of sheet and rows with custom settings will be yield.
  ## If restrict is true, only rows in dimension of sheet will be yield.
  var all: HashSet[int]
  for r in xs.first.row..xs.last.row:
    if r >= 0: all.incl r
  for r in xs.rows.keys: all.incl r

  for r in all.sortedItems:
    if restrict and r notin xs.first.row..xs.last.row:
      continue

    yield xs.row(r)

iterator cols*(xs: XlSheet, restrict = false): XlCol =
  ## Iterate over columns in a sheet. Yield XlCol objects.
  ## Both columns in dimension of sheet and columns with custom settings will be yield.
  ## If restrict is true, only columns in dimension of sheet will be yield.
  var all: HashSet[int]
  for c in xs.first.col..xs.last.col:
    if c >= 0: all.incl c
  for c in xs.cols.keys: all.incl c

  for c in all.sortedItems:
    if restrict and c notin xs.first.col..xs.last.col:
      continue

    yield xs.col(c)

iterator rows*(xr: XlRange): XlRange =
  ## Iterate over rows in a range.
  ## Notice, yield XlRange object instead of XlRow object.
  for row in xr.first.row..xr.last.row:
    yield XlRange(sheet: xr.sheet,
      first: (row, xr.first.col),
      last: (row, xr.last.col)
    )

iterator cols*(xr: XlRange): XlRange =
  ## Iterate over columns in a range.
  ## Notice, yield XlRange object instead of XlRow object.
  for col in xr.first.col..xr.last.col:
    yield XlRange(sheet: xr.sheet,
      first: (xr.first.row, col),
      last: (xr.last.row, col)
    )

proc collection*(xs: XlSheet, empty = false): XlCollection =
  ## Return a collection of cells from a sheet.
  ## If `empty` == false, the returned collection contains all cells in the sheet.
  result = XlCollection(sheet: xs)
  if not empty:
    for rc in xs.cells.keys:
      result.rcs.incl rc

proc collecttion*(xr: XlRange): XlCollection {.inline.} =
  ## Return a collection of cells from a range.
  ## The returned collection will contains all nonempty cells in the range.
  result = XlCollection(sheet: xr.sheet)
  for rc in xr.sheet.cells.keys:
    if rc in xr:
      result.rcs.incl rc

proc add*(xco: XlCollection, c: XlRC|XlRCRC|XlCell|XlRange|string, includeEmpty=true) =
  ## Add a cell or a range into a collection.
  ## If includeEmpty is false, only cells that actually in working area will be added.
  when c is XlRC:
    if includeEmpty or c in xco.sheet:
      xco.rcs.incl c

  elif c is XlCell:
    # assert c.sheet == xco.sheet
    if includeEmpty or c in xco.sheet:
      xco.rcs.incl c.rc

  else:
    var first, last: XlRC
    when c is XlRange:
      (first, last) = (c.first, c.last)

    elif c is string:
      (first, last) = c.rcrc

    elif c is XlRCRC:
      (first, last) = c

    for row in first.row..last.row:
      for col in first.col..last.col:
        if includeEmpty or (row, col) in xco.sheet.cells:
          xco.rcs.incl (row, col)

proc delete*(xco: XlCollection, c: XlRC|XlRCRC|XlCell|XlRange|string) =
  ## Delete a cell or a range from a collection.
  when c is XlRC:
    xco.rcs.excl c

  elif c is XlCell:
    xco.rcs.excl c.rc

  else:
    var first, last: XlRC
    when c is XlRange:
      (first, last) = (c.first, c.last)

    elif c is string:
      (first, last) = c.rcrc

    elif c is XlRCRC:
      (first, last) = c

    for row in first.row..last.row:
      for col in first.col..last.col:
        xco.rcs.excl (row, col)

proc clear*(xco: XlCollection) {.inline.} =
  ## Clear a collection.
  xco.rcs.clear()

proc count*(xco: XlCollection): int {.inline.} =
  ## Return count of cells in a collection.
  return xco.rcs.len

proc len*(xco: XlCollection): int {.inline.} =
  ## Return count of cells in a collection.
  return xco.rcs.len

proc turn*(xr: XlRange): XlRange =
  ## Turn iteration direction of a range.
  ## It is vertical (row first) by default.
  return XlRange(
    sheet: xr.sheet,
    first: xr.first,
    last: xr.last,
    turned: not xr.turned)

iterator items*(xr: XlRange): XlCell =
  ## Iterate over cells in a range.
  ## The direction (horizontal or vertical) can be change by `turn` proc.
  # Caution: the body of a for loop over an inline iterator is inlined
  # into each yield statement appearing in the iterator code, so ideally
  # the code should be refactored to contain a single yield when possible
  # to avoid code bloat.
  var
    level1 = xr.first.row..xr.last.row
    level2 = xr.first.col..xr.last.col

  if xr.turned:
    swap level1, level2

  for i1 in level1:
    for i2 in level2:
      var rc = if xr.turned: (i2, i1) else: (i1, i2)
      yield xr.sheet.cell(rc)

iterator items*(xco: XlCollection): XlCell =
  ## Iterate over cells in a collection.
  for rc in xco.rcs.sortedItems:
    yield xco.sheet.cell(rc)

iterator items*(xc: XlCell): XlCell =
  ## Yield a cell itself. Nop, for convenient only.
  yield xc

proc movecopy(xr: XlRange, to: XlRC, isMove=true) =
  if xr.isEmpty or xr.first == to:
    return

  let
    xs = xr.sheet
    diffr = to.row - xr.first.row
    diffc = to.col - xr.first.col

  defer:
    xs.update()

  var rcs: seq[XlRC]
  for rc in xs.cells.keys:
    if rc in xr:
      rcs.add rc

  # to deal with cell overlap
  if diffr <= 0 and diffc <= 0:
    rcs.sort do (a, b: XlRC) -> int:
      result = cmp(a.row, b.row)
      if result == 0: result = cmp(a.col, b.col)

  elif diffr <= 0 and diffc > 0:
    rcs.sort do (a, b: XlRC) -> int:
      result = cmp(a.row, b.row)
      if result == 0: result = cmp(b.col, a.col)

  elif diffr > 0 and diffc <= 0:
    rcs.sort do (a, b: XlRC) -> int:
      result = cmp(b.row, a.row)
      if result == 0: result = cmp(a.col, b.col)

  else:
    rcs.sort do (a, b: XlRC) -> int:
      result = cmp(b.row, a.row)
      if result == 0: result = cmp(b.col, a.col)

  for rc in rcs:
    let newrc = (rc.row + diffr, rc.col + diffc)
    var cell = xs.cells.getOrDefault(rc, nil)
    if cell != nil:
      if isMove:
        xs.cells.del rc
      else:
        cell = cell.dup

      cell.rc = newrc
      xs.cells[newrc] = cell

  var rcrcs: seq[(XlRC, XlRC)]
  for rcrc in xs.merges:
    let
      firstIn = rcrc[0] in xr
      lastIn = rcrc[1] in xr

    if firstIn and lastIn:
      rcrcs.add rcrc

    elif (firstIn and not lastIn) or (not firstIn and lastIn):
      raise newException(XlError,
        "Invalid operation for merged range: " & $ rcrc.name)

  for rcrc in rcrcs:
    xs.merges.excl rcrc
    xs.merges.incl ((rcrc[0].row + diffr, rcrc[0].col + diffc),
      (rcrc[1].row + diffr, rcrc[1].col + diffc))

proc move*(xr: XlRange, to: XlRC|string) {.inline.} =
  ## Move a range.
  ## Notice: cell reference in formulas and tables won't be chnaged.
  when to is string:
    movecopy(xr, to.rc, isMove=true)
  else:
    movecopy(xr, to, isMove=true)

proc copy*(xr: XlRange, to: XlRC|string) {.inline.} =
  ## Copy a range.
  ## Notice: cell reference in formulas and tables won't be chnaged.
  when to is string:
    movecopy(xr, to.rc, isMove=false)
  else:
    movecopy(xr, to, isMove=false)

proc insertRow*(xs: XlSheet, at: int, count = 1) =
  ## Insert empty rows into a sheet.
  ## Notice: cell reference in formulas and tables won't be chnaged.
  if count < 1:
    raise newException(XlError, "invalid count")

  XlRange(sheet: xs,
    first: (at, xs.first.col),
    last: (xs.last.row, xs.last.col)
  ).movecopy((at + count, xs.first.col), true)

  var rows: seq[int]
  for r in xs.rows.keys:
    if r >= at:
      rows.add r

  rows.sort()
  for i in countdown(rows.high, 0):
    let r = rows[i]
    var row = xs.rows[r]
    row.row = r
    xs.rows[r + count] = row
    xs.rows.del r

proc deleteRow*(xs: XlSheet, at: int, count = 1) =
  ## Delete rows from a sheet.
  ## Notice: cell reference in formulas and tables won't be chnaged.
  if count < 1:
    raise newException(XlError, "invalid count")

  XlRange(sheet: xs,
    first: (at + count, xs.first.col),
    last: (xs.last.row, xs.last.col)
  ).movecopy((at, xs.first.col), true)

  var rows: seq[int]
  for r in xs.rows.keys:
    if r >= at:
      rows.add r

  rows.sort()
  for r in rows:
    if r >= at + count:
      var row = xs.rows[r]
      row.row = r
      xs.rows[r - count] = row
    xs.rows.del r

proc insertCol*(xs: XlSheet, at: int, count = 1) =
  ## Insert empty columns into a sheet.
  ## Notice: cell reference in formulas and tables won't be chnaged.
  if count < 1:
    raise newException(XlError, "invalid count")

  XlRange(sheet: xs,
    first: (xs.first.row, at),
    last: (xs.last.row, xs.last.col)
  ).movecopy((xs.first.row, at + count), true)

  var cols: seq[int]
  for c in xs.cols.keys:
    if c >= at:
      cols.add c

  cols.sort()
  for i in countdown(cols.high, 0):
    let c = cols[i]
    var col = xs.cols[c]
    col.col = c
    xs.cols[c + count] = col
    xs.cols.del c

proc deleteCol*(xs: XlSheet, at: int, count = 1) =
  ## Delete columns from a sheet.
  ## Notice: cell reference in formulas and tables won't be chnaged.
  if count < 1:
    raise newException(XlError, "invalid count")

  XlRange(sheet: xs,
    first: (xs.first.row, at + count),
    last: (xs.last.row, xs.last.col)
  ).movecopy((xs.first.row, at), true)

  var cols: seq[int]
  for c in xs.cols.keys:
    if c >= at:
      cols.add c

  cols.sort()
  for c in cols:
    if c >= at + count:
      var col = xs.cols[c]
      col.col = c
      xs.cols[c - count] = col
    xs.cols.del c

proc unmerge*(x: XlRange|XlCollection|XlCell) =
  ## Unmege a range, or unmerge any range or collection include a cell.
  for xc in x:
    for rcrc in xc.sheet.merges:
      if xc.rc in rcrc:
        xc.sheet.merges.excl rcrc
        break

proc merge*(xr: XlRange) {.inline.} =
  ## Merge a range.
  xr.unmerge
  xr.sheet.merges.incl (xr.first, xr.last)

proc editCheck(x: XlCell|XlRow|XlCol) {.inline.} =
  when not defined(release):
    when x is XlCell:
      if x.readonly:
        raise newException(XlError, "Cell is readonly")

      if x.rc notin x.sheet.cells:
        raise newException(XlError, "Cell is not in sheet, had been deleted?")
    elif x is XlRow:
      if x.row notin x.sheet.rows:
        raise newException(XlError, "Row is not in sheet, had been deleted?")
    elif x is XlCol:
      if x.col notin x.sheet.cols:
        raise newException(XlError, "Column is not in sheet, had been deleted?")

proc height*(x: XlRow): float {.inline.} =
  ## Return height of a row. -1 means default.
  return x.height

proc `height=`*(x: XlRow, height: float) {.inline.} =
  ## Set height of a row.
  ## Assign -1 to reset to default.
  x.editCheck()
  x.height = height

proc width*(x: XlCol): float {.inline.} =
  ## Return width of a column. -1 means default.
  return x.width

proc `width=`*(x: XlCol, width: float) {.inline.} =
  ## Set width of a column.
  ## Assign -1 to reset to default.
  x.editCheck()
  x.width = width

proc hidden*(x: XlRow|XlCol|XlSheet): bool {.inline.} =
  ## Return if a row, column, or sheet is hidden.
  return x.hidden

proc `hidden=`*(x: XlRow|XlCol|XlSheet, hidden: bool) {.inline.} =
  ## Set a row, column, or sheet is hidden or not.
  when x is XlRow|XlCol:
    x.editCheck()
  x.hidden = hidden

proc isNumber*(xc: XlCell): bool {.inline.} =
  ## Return the cell is number or not.
  return xc.ct == ctNumber

template write[T](x: XlRange, data: openArray[T], setter: untyped): untyped =
  var index = 0
  for xc in x:
    if index >= data.len: break
    setter(xc, data[index])
    index.inc

proc value*(xc: XlCell): string {.inline.} =
  ## Return the text value of a cell.
  return xc.value

proc `value=`*(x: XlRange|XlCollection|XlCell, val: string|SomeNumber|DateTime) =
  ## Set the text, number, or date value of a cell, or all cells in range or collection.
  for xc in x:
    xc.editCheck()
    xc.riches.setLen(0)
    when val is string:
      xc.value = val
      xc.ct = ctSharedString
    elif val is DateTime:
      xc.value = $(inSeconds(val - dateTime(1900, mJan, 1)).float / 86400)
      xc.ct = ctNumber
    else:
      xc.value = $val
      xc.ct = ctNumber

proc `value=`*[T: string|SomeNumber|DateTime](x: XlRange, data: openArray[T]) =
  ## Put a text, number, or date array into range. Extra data will be discarded.
  # cannot use write template, compiler's bug?
  var index = 0
  for xc in x:
    if index >= data.len: break
    `value=`(xc, data[index])
    index.inc

template `value=`*(x: XlRange, tup: tuple) =
  ## Put a tuple contains text, number, or date into range. Extra data will be discarded.
  var cells: seq[XlCell]
  for xc in x:
    cells.add xc

  var index = 0
  for value in tup.fields:
    if index >= cells.len: break
    `value=`(cells[index], value)
    index.inc

proc number*(xc: XlCell): float {.inline.} =
  ## Return the number value of a cell.
  ## Raise `ValueError` on fail.
  return parseFloat(xc.value)

proc `number=`*(x: XlRange|XlCollection|XlCell, val: SomeNumber) {.inline.} =
  ## Set the number value of a cell, or all cells in range or collection.
  `value=`(x, val)

proc `number=`*[T: SomeNumber](x: XlRange, data: openArray[T]) {.inline.} =
  ## Put a number array into range. Extra data will be discarded.
  `value=`(x, data)

proc date*(xc: XlCell): DateTime =
  ## Assume the number of cell is a date and return it.
  ## Using 1900 compatibility date system: a serial number that represents the number of days
  ## elapsed since December, 31st 1899. It also (wrongly) counts 1900 as a leap year.
  result = dateTime(1899, mDec, 31) 
  var (nDays, dayFraction) = splitDecimal(xc.number)
  if nDays >= 61:
    nDays -= 1
  result += int(nDays).days + int(round(dayFraction * 86400, 0)).seconds

proc `date=`*(x: XlRange|XlCollection|XlCell, date: DateTime) =
  ## Set the date of a cell, or all cells in range or collection.
  ## Using 1900 date system: a serial number that represents the number of days
  ## elapsed since January 1, 1900.
  `value=`(x, date)

proc `date=`*(x: XlRange, dates: openArray[DateTime]) =
  ## Put a date array into range. Extra data will be discarded.
  `value=`(x, dates)

proc riches*(xc: XlCell): XlRiches {.inline.} =
  ## Return XlRiches type of a cell, it represent a rich string.
  return xc.riches

proc `riches=`*(x: XlRange|XlCollection|XlCell, riches: openArray[XlRich]) {.inline.} =
  ## Set the rich string of a cell, or all cells in range or collection.
  ## For example:
  ##
  ## .. code-block:: Nim
  ##   sheet.cell("A1").riches = {
  ##    "big": XlFont(size: 16.0),
  ##    "small": XlFont(size: 8.0)
  ##   }
  for xc in x:
    xc.editCheck()
    xc.riches = @riches
    xc.value = xc.riches.value
    xc.ct = ctInlineStr

proc formula*(xc: XlCell): string {.inline.} =
  ## Return formula of a cell.
  return xc.formula

proc `formula=`*(x: XlRange|XlCollection|XlCell, f: string) {.inline.} =
  ## Set formula of a cell, or all cells in range or collection.
  ## Notice: prefix "=" is not needed.
  for xc in x:
    xc.editCheck()
    xc.formula = f

proc `formula=`*(x: XlRange, data: openArray[string]) =
  ## Put a formula array into range. Extra formulas will be discarded.
  write(x, data, `formula=`)

proc hyperlink*(xc: XlCell): string {.inline.} =
  ## Return hyperlink of a cell.
  return xc.hyperlink

proc `hyperlink=`*(x: XlRange|XlCollection|XlCell, h: string) {.inline.} =
  ## Set hyperlink to a cell, or all cells in range or collection.
  for xc in x:
    xc.editCheck()
    xc.hyperlink = h

proc `hyperlink=`*(x: XlRange, data: openArray[string]) =
  ## Put a hyperlink array into range. Extra hyperlinks will be discarded.
  write(x, data, `hyperlink=`)

proc fetchStyle[T](x: XlStylable, result: var T) {.inline.} =

  template fetch(sym: untyped, parseFn: untyped): untyped =
    if not x.change.isNil and x.change.sym.isSome:
        result = x.change.sym.get

    elif x.style >= 0:
      result = x.sheet.workbook.sharedStyles.xStyles.parseFn(x.style)

  when T is XlNumFmt: fetch(numFmt, parseNumFmt)
  elif T is XlFont: fetch(font, parseFont)
  elif T is XlFill: fetch(fill, parseFill)
  elif T is XlBorder: fetch(border, parseBorder)
  elif T is XlAlignment: fetch(alignment, parseAlignment)
  elif T is XlProtection: fetch(protection, parseProtection)

proc numFmt*(x: XlStylable): XlNumFmt {.inline.} =
  ## Return XlNumFmt object of a XlCell, XlRow, or XlCol.
  ## The XlNumFmt object represent the number format style.
  fetchStyle(x, result)

proc font*(x: XlStylable): XlFont {.inline.} =
  ## Return XlFont object of a XlCell, XlRow, or XlCol.
  ## The XlFont object represent the font style.
  fetchStyle(x, result)

proc fill*(x: XlStylable): XlFill {.inline.} =
  ## Return XlFill object of a XlCell, XlRow, or XlCol.
  ## The XlFill object represent the fill style.
  fetchStyle(x, result)

proc border*(x: XlStylable): XlBorder {.inline.} =
  ## Return XlBorder object of a XlCell, XlRow, or XlCol.
  ## The XlBorder object represent the border style.
  fetchStyle(x, result)

proc alignment*(x: XlStylable): XlAlignment {.inline.} =
  ## Return XlAlignment object of a XlCell, XlRow, or XlCol.
  ## The XlAlignment object represent the alignment style.
  fetchStyle(x, result)

proc protection*(x: XlStylable): XlProtection {.inline.} =
  ## Return XlProtection object of a XlCell, XlRow, or XlCol.
  ## The XlProtection object represent the protection setting.
  fetchStyle(x, result)

proc style*(x: XlStylable): XlStyle =
  ## Return XlStyle object of a XlCell, XlRow, or XlCol.
  ## The XlStyle object represent all the styles.
  result.numFmt = x.numFmt
  result.font = x.font
  result.fill = x.fill
  result.border = x.border
  result.alignment = x.alignment
  result.protection = x.protection

proc applyStyle[T](x: XlStylable|XlRange|XlCollection, s: T) =

  template apply(sym: untyped): untyped =
    when x is XlStylable:
      x.editCheck()
      if x.change.isNil: x.change = XlChange()
      x.change.sym = some s
      validate(x.change.sym.get, error=true)

    else:
      for xc in x:
        applyStyle(xc, s)

  when T is XlNumFmt: apply(numFmt)
  elif T is XlFont: apply(font)
  elif T is XlFill: apply(fill)
  elif T is XlBorder: apply(border)
  elif T is XlAlignment: apply(alignment)
  elif T is XlProtection: apply(protection)

  # apply style to cells in row/col
  when x is XlStylable:
    if x is XlRow:
      let row = cast[XlRow](x)
      for cell in x.sheet.cells.values:
        if cell.rc.row == row.row:
          applyStyle(cell, s)

    elif x is XlCol:
      let col = cast[XlCol](x)
      for cell in x.sheet.cells.values:
        if cell.rc.col == col.col:
          applyStyle(cell, s)

proc `numFmt=`*(x: XlStylable|XlRange|XlCollection, numFmt: XlNumFmt) {.inline.} =
  ## Set number format of a cell, row, or column; or all cells in range or collection.
  ## Assign default(XlNumFmt) to reset to default.
  x.applyStyle(numFmt)

proc `numFmt=`*(x: XlRange, data: openArray[XlNumFmt]) =
  ## Put a number format array into range. Extra number formats will be discarded.
  write(x, data, `numFmt=`)

proc `font=`*(x: XlStylable|XlRange|XlCollection, font: XlFont) {.inline.} =
  ## Set font of a cell, row, or column; or all cells in range or collection.
  ## Assign default(XlFont) to reset to default.
  x.applyStyle(font)

proc `font=`*(x: XlRange, data: openArray[XlFont]) =
  ## Put a font array into range. Extra font will be discarded.
  write(x, data, `font=`)

proc `fill=`*(x: XlStylable|XlRange|XlCollection, fill: XlFill) {.inline.} =
  ## Set fill of a cell, row, or column; or all cells in range or collection.
  ## Assign default(XlFill) to reset to default.
  x.applyStyle(fill)

proc `fill=`*(x: XlRange, data: openArray[XlFill]) =
  ## Put a fill array into range. Extra fill will be discarded.
  write(x, data, `fill=`)

proc `border=`*(x: XlStylable|XlRange|XlCollection, border: XlBorder) {.inline.} =
  ## Set border of a cell, row, or column; or all cells in range or collection.
  ## Assign default(XlBorder) to reset to default.
  x.applyStyle(border)

proc `border=`*(x: XlRange, data: openArray[XlBorder]) =
  ## Put a border array into range. Extra border will be discarded.
  write(x, data, `border=`)

proc `alignment=`*(x: XlStylable|XlRange|XlCollection, alignment: XlAlignment) {.inline.} =
  ## Set alignment of a cell, row, or column; or all cells in range or collection.
  ## Assign default(XlAlignment) to reset to default.
  x.applyStyle(alignment)

proc `alignment=`*(x: XlRange, data: openArray[XlAlignment]) =
  ## Put a alignment array into range. Extra alignment will be discarded.
  write(x, data, `alignment=`)

proc `protection=`*(x: XlStylable|XlRange|XlCollection, protection: XlProtection) {.inline.} =
  ## Set protection of a cell, row, or column; or all cells in range or collection.
  ## Assign default(XlProtection) to reset to default.
  x.applyStyle(protection)

proc `protection=`*(x: XlRange, data: openArray[XlProtection]) =
  ## Put a protection array into range. Extra protection will be discarded.
  write(x, data, `protection=`)

proc `style=`*(x: XlStylable|XlRange|XlCollection, style: XlStyle) =
  ## Set style of a cell, row, or column; or all cells in range or collection.
  ## Assign default(XlStyle) to reset to default.
  when x is XlStylable:
    x.editCheck()
    if style == default(XlStyle):
      # reset all style
      if x.change.isNil: x.change = XlChange()
      x.change[] = default(XlChange()[].type)
    else:
      x.numFmt = style.numFmt
      x.font = style.font
      x.fill = style.fill
      x.border = style.border
      x.alignment = style.alignment
      x.protection = style.protection

  else:
    for xc in x:
      `style=`(xc, style)

proc `style=`*(x: XlRange, data: openArray[XlStyle]) =
  ## Put a style array into range. Extra style will be discarded.
  write(x, data, `style=`)

proc `top=`*(x: XlCell|XlRange|XlCollection, side: XlSide) =
  ## Set top border of a cell, or all cells in range or collection.
  for c in x:
    var border = c.border
    border.top = some side
    c.border = border

proc `bottom=`*(x: XlCell|XlRange|XlCollection, side: XlSide) =
  ## Set bottom border of a cell, or all cells in range or collection.
  for c in x:
    var border = c.border
    border.bottom = some side
    c.border = border

proc `left=`*(x: XlCell|XlRange|XlCollection, side: XlSide) =
  ## Set left border of a cell, or all cells in range or collection.
  for c in x:
    var border = c.border
    border.left = some side
    c.border = border

proc `right=`*(x: XlCell|XlRange|XlCollection, side: XlSide) =
  ## Set right border of a cell, or all cells in range or collection.
  for c in x:
    var border = c.border
    border.right = some side
    c.border = border

proc `outline=`*(x: XlRange, side: XlSide) =
  ## Set outline border of a range.
  x.row(x.first.row).top = side
  x.row(x.last.row).bottom = side
  x.col(x.first.col).left = side
  x.col(x.last.col).right = side

proc `horizontal=`*(x: XlRange, side: XlSide) =
  ## Set horizontal inside border of a range.
  for i in x.first.row..(x.last.row - 1):
    x.row(i).bottom = side

  for i in (x.first.row + 1)..x.last.row:
    x.row(i).top = side

proc `vertical=`*(x: XlRange, side: XlSide) =
  ## Set vertical inside border of a range.
  for i in x.first.col..(x.last.col - 1):
    x.col(i).right = side

  for i in (x.first.col + 1)..x.last.col:
    x.col(i).left = side
