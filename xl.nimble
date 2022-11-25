#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#                  Copyright (c) 2022 Ward
#
#====================================================================

# Package
version       = "1.0.1"
author        = "Ward"
description   = "Xl - Open XML Spreadsheet (Excel) Library for Nim"
license       = "MIT"
skipDirs      = @["examples", "htmldocs"]

# Dependencies
requires "nim >= 1.6.0"
requires "zippy >= 0.10.4"

# Examples
task example, "Build all the examples":
  cd "examples"
  defer: cd ".."
  exec "nim r cell_referencing.nim"
  exec "nim r demo.nim"
  exec "nim r doc_properties.nim"
  exec "nim r hello_world.nim"
  exec "nim r hyperlink.nim"
  exec "nim r merge_rich_string.nim"
  exec "nim r protection.nim"
  exec "nim r skyscrapers.nim"
  exec "nim r styles.nim"
  exec "nim r template.nim"
  exec "nim r tutorial1.nim"
  exec "nim r tutorial2.nim"
  exec "nim r tutorial3.nim"

# Clean
task clean, "Delete all xlsx files":
  exec "cmd /c IF EXIST examples\\*.xlsx del examples\\*.xlsx"
