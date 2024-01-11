import ../../xl
import std/times

let workbook = load("tests/dates/dates.xlsx")
let firstSheet = workbook.sheet(0)

# Dates without time information (A1:A3)

doAssert firstSheet.cell("A1").date == dateTime(2022, Month.mFeb, 12, 0, 0, 0, 0, local())
doAssert firstSheet.cell("A2").date == dateTime(1945, Month.mDec, 25, 0, 0, 0, 0, local())
doAssert firstSheet.cell("A3").date == dateTime(1995, Month.mApr, 01, 0, 0, 0, 0, local())
doAssert firstSheet.cell("A4").date == dateTime(1900, Month.mJan, 11, 0, 0, 0, 0, local())
doAssert firstSheet.cell("A5").date == dateTime(1900, Month.mMar, 03, 0, 0, 0, 0, local())


# Datetime cells (A4:A6)
echo $firstSheet.cell("A7").date, $firstSheet.cell("A7").value
doAssert firstSheet.cell("A7").date == dateTime(1900, Month.mMar, 03, 13, 0, 0, 0, local())
