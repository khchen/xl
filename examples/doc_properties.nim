#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#               Copyright (c) Chen Kai-Hung, Ward
#
#====================================================================

# Example source: https://xlsxwriter.readthedocs.io/example_doc_properties.html

import xl, times

var workbook = newWorkbook()
var sheet = workbook.add("Properties")

workbook.properties = XlProperties(
  title: "This is an example spreadsheet",
  subject: "With document properties",
  creator: "Xl",
  keywords: "Sample, Example, Properties",
  description: "Created with Nim Xl",
  lastModifiedBy: "Xl",
  created: $now(),
  modified: $now(),
  category: "Example spreadsheets",
  contentStatus: "Nim",
  application: "Microsoft Excel",
  manager: "Who?",
  company: "https://nim-lang.org/"
)

sheet.col(0).width = 70
sheet.cell("A1").value = "Select 'Workbook Properties' to see properties."

workbook.save("doc_properties.xlsx")
