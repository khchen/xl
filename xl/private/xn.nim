#====================================================================
#
#         Xl - Open XML Spreadsheet (Excel) Library for Nim
#                  Copyright (c) 2022 Ward
#
#====================================================================

# XlNode is XmlNode to deal with namespace and children.
# The code copy from std/xmltree and std/xmlparser

import strtabs, strutils, streams, parsexml, tables
export parsexml

type
  XlNodeKind* = enum
    xnText,
    xnElement,
    xnCData,
    xnEntity

  XlAttributes* = StringTableRef

  XlNsName* = object
    ns*: string
    name*: string

  XlNode* = ref object
    case kind: XlNodeKind
    of xnText, xnCData, xnEntity:
      fText: string
    of xnElement:
      fTag: string
      s: seq[XlNode]
      attrs*: XlAttributes
      parent {.cursor.}: XlNode
      childrenI: seq[int]
      childrenOk: bool
      nsmap: XlAttributes
      nstag: XlNsName

const
  xlHeader* = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n"

proc newXlNode(kind: XlNodeKind): XlNode =
  result = XlNode(kind: kind)

proc newElement*(tag: string): XlNode =
  result = newXlNode(xnElement)
  result.fTag = tag
  result.s = @[]

proc newText*(text: string): XlNode =
  result = newXlNode(xnText)
  result.fText = text

proc newCData*(cdata: string): XlNode =
  result = newXlNode(xnCData)
  result.fText = cdata

proc newEntity*(entity: string): XlNode =
  result = newXlNode(xnEntity)
  result.fText = entity

proc innerText*(n: XlNode): string =
  proc worker(res: var string, n: XlNode) =
    case n.kind
    of xnText, xnEntity:
      res.add(n.fText)
    of xnElement:
      for sub in n.s:
        worker(res, sub)
    else:
      discard

  result = ""
  worker(result, n)

proc rawLen*(n: XlNode): int {.inline.} =
  if n.kind == xnElement:
    return len(n.s)

iterator items(n: XlNode): XlNode {.inline.} =
  # remove * to avoid unexpected usage
  # use iterator children instead
  assert n.kind == xnElement
  for i in 0 .. n.s.len-1:
    yield n.s[i]

iterator pairs(n: XlNode): (int, XlNode) {.inline.} =
  # remove * to avoid unexpected usage
  # use iterator childrenPair instead
  assert n.kind == xnElement
  for i in 0 .. n.s.len-1:
    yield (i, n.s[i])

proc toXlAttributes*(keyValuePairs: varargs[tuple[key, val: string]]): XlAttributes =
  newStringTable(keyValuePairs)

proc attr*(n: XlNode, name: string): string =
  assert n.kind == xnElement
  if n.attrs == nil: return ""
  return n.attrs.getOrDefault(name)

proc `[]`*(n: XlNode, key: string): string {.inline.} =
  return n.attr(key)

proc `[]=`*(n: XlNode, key: string, value: string) =
  assert n.kind == xnElement
  if n.attrs == nil:
    n.attrs = {key: value}.toXlAttributes
  else:
    n.attrs[key] = value

proc `{}=`(n: XlNode, prefix: string, ns: string) =
  if n.nsmap == nil:
    n.nsmap = {prefix: ns}.toXlAttributes
  else:
    n.nsmap[prefix] = ns

proc addEscaped*(result: var string, s: string) =
  ## The same as `result.add(escape(s)) <#escape,string>`_, but more efficient.
  for c in items(s):
    case c
    of '<': result.add("&lt;")
    of '>': result.add("&gt;")
    of '&': result.add("&amp;")
    of '"': result.add("&quot;")
    of '\'': result.add("&apos;")
    else: result.add(c)

proc addIndent(result: var string, indent: int, addNewLines: bool) =
  if addNewLines:
    result.add("\n")
  for i in 1 .. indent:
    result.add(' ')

proc add*(result: var string, n: XlNode, indent = 0, indWidth = 2,
          addNewLines = true) =
  proc noWhitespace(n: XlNode): bool =
    for i in 0 ..< n.rawLen:
      if n.s[i].kind in {xnText, xnEntity}: return true

  proc addEscapedAttr(result: var string, s: string) =
    # `addEscaped` alternative with less escaped characters.
    # Only to be used for escaping attribute values enclosed in double quotes!
    for c in items(s):
      case c
      of '<': result.add("&lt;")
      of '>': result.add("&gt;")
      of '&': result.add("&amp;")
      of '"': result.add("&quot;")
      else: result.add(c)

  if n == nil: return

  case n.kind
  of xnElement:
    if indent > 0:
      result.addIndent(indent, addNewLines)

    let
      addNewLines = if n.noWhitespace():
                      false
                    else:
                      addNewLines

    result.add('<')
    result.add(n.fTag)
    if not isNil(n.attrs):
      for key, val in pairs(n.attrs):
        result.add(' ')
        result.add(key)
        result.add("=\"")
        result.addEscapedAttr(val)
        result.add('"')

    if n.rawLen == 0:
      result.add(" />")
      return

    let
      indentNext = if n.noWhitespace():
                     indent
                   else:
                     indent+indWidth
    result.add('>')
    for i in 0 ..< n.rawLen:
      result.add(n.s[i], indentNext, indWidth, addNewLines)

    if not n.noWhitespace():
      result.addIndent(indent, addNewLines)

    result.add("</")
    result.add(n.fTag)
    result.add(">")
  of xnText:
    result.addEscaped(n.fText)
  of xnCData:
    result.add("<![CDATA[")
    result.add(n.fText)
    result.add("]]>")
  of xnEntity:
    result.add('&')
    result.add(n.fText)
    result.add(';')

proc `$`*(n: XlNode): string =
  result.add(n)

proc rawChild*(n: XlNode, name: string): XlNode =
  # rename child to rawChild to return child without namespace handling
  # assert n.kind == xnElement
  for i in items(n):
    if i.kind == xnElement:
      if i.fTag == name:
        return i

type
  XlParserError* = object of ValueError
    errors*: seq[string]

proc raiseInvalidXml(errors: seq[string]) =
  var e: ref XlParserError
  new(e)
  e.msg = errors[0]
  e.errors = errors
  raise e

proc addNode(father, son: XlNode) =
  if son != nil:
    add(father.s, son)
    if son.kind == xnElement:
      son.parent = father

proc parse*(x: var XmlParser, errors: var seq[string]): XlNode {.gcsafe.}

proc untilElementEnd(x: var XmlParser, result: XlNode, errors: var seq[string]) =
  while true:
    case x.kind
    of xmlElementEnd:
      if x.elementName == result.fTag:
        next(x)
      else:
        errors.add(errorMsg(x, "</" & result.fTag & "> expected"))
        # do not skip it here!
      break
    of xmlEof:
      errors.add(errorMsg(x, "</" & result.fTag & "> expected"))
      break
    else:
      result.addNode(parse(x, errors))

proc parse*(x: var XmlParser, errors: var seq[string]): XlNode =
  case x.kind
  of xmlCharData, xmlWhitespace:
    result = newText(x.charData)
    next(x)
  of xmlPI, xmlSpecial, xmlComment:
    # we just ignore processing instructions for now
    next(x)
  of xmlError:
    errors.add(errorMsg(x))
    next(x)
  of xmlElementStart: ## ``<elem>``
    result = newElement(x.elementName)
    next(x)
    untilElementEnd(x, result, errors)
  of xmlElementEnd:
    errors.add(errorMsg(x, "unexpected ending tag: " & x.elementName))
  of xmlElementOpen:
    result = newElement(x.elementName)
    next(x)
    while true:
      case x.kind
      of xmlAttribute:
        if x.attrKey == "xmlns":
          result{""} = x.attrValue
        elif x.attrKey.startsWith("xmlns:"):
          result{x.attrKey[6..^1]} = x.attrValue
        result[x.attrKey] = x.attrValue
        next(x)
      of xmlElementClose:
        next(x)
        break
      of xmlError:
        errors.add(errorMsg(x))
        next(x)
        break
      else:
        errors.add(errorMsg(x, "'>' expected"))
        next(x)
        break
    untilElementEnd(x, result, errors)
  of xmlAttribute, xmlElementClose:
    errors.add(errorMsg(x, "<some_tag> expected"))
    next(x)
  of xmlCData:
    result = newCData(x.charData)
    next(x)
  of xmlEntity:
    ## &entity;
    result = newEntity(x.entityName)
    next(x)
  of xmlEof: discard

proc parseXml*(s: Stream, filename: string,
               errors: var seq[string], options: set[XmlParseOption] = {reportComments}): XlNode =
  ## Parses the XML from stream ``s`` and returns a ``XlNode``. Every
  ## occurred parsing error is added to the ``errors`` sequence.
  var x: XmlParser
  open(x, s, filename, options)
  while true:
    x.next()
    case x.kind
    of xmlElementOpen, xmlElementStart:
      result = parse(x, errors)
      break
    of xmlComment, xmlWhitespace, xmlSpecial, xmlPI: discard # just skip it
    of xmlError:
      errors.add(errorMsg(x))
    else:
      errors.add(errorMsg(x, "<some_tag> expected"))
      break
  close(x)

proc parseXml*(s: Stream, options: set[XmlParseOption] = {reportComments}): XlNode =
  ## Parses the XML from stream ``s`` and returns a ``XlNode``. All parsing
  ## errors are turned into an ``XmlError`` exception.
  var errors: seq[string] = @[]
  result = parseXml(s, "memory_xml", errors, options)
  if errors.len > 0: raiseInvalidXml(errors)

proc parseXml*(str: string, options: set[XmlParseOption] = {reportComments}): XlNode =
  ## Parses the XML from string ``str`` and returns a ``XlNode``. All parsing
  ## errors are turned into an ``XmlError`` exception.
  parseXml(newStringStream(str), options)

proc `$$`*(n: XlNode): string =
  # return ugly xml
  result.add(n, indent=0, indWidth=0, addNewLines=false)

template `\`*(a: string, b: string): XlNsName =
  XlNsName(ns: a, name: b)

proc parseXn*(xml: string): XlNode =
  # default xml to XlNode parser
  result = parseXml(xml, {reportWhitespace})

proc addNameSpace*(x: XlNode, ns: string, prefix = "") =
  if prefix != "":
    x["xmlns:" & prefix] = ns
  else:
    x["xmlns"] = ns
  x{prefix} = ns

proc prepareChildren(x: XlNode) =
  # collect xnElement nodes
  assert x.kind == xnElement
  if not x.childrenOk:
    x.childrenI.setLen(0)
    for i, n in x:
      if n.kind == xnElement:
        x.childrenI.add i
    x.childrenOk = true

iterator children*(x: XlNode): XlNode =
  x.prepareChildren()
  for i in x.childrenI:
    yield x.s[i]

iterator childrenPair*(x: XlNode): (int, XlNode) =
  x.prepareChildren()
  for index, i in x.childrenI:
    yield (index, x.s[i])

proc `[]`*(x: XlNode, i: int): XlNode {.inline.} =
  x.prepareChildren()
  return x.s[x.childrenI[i]]

proc count*(x: XlNode): int {.inline.} =
  # not provide `len` to distinguish rawLen and count
  x.prepareChildren()
  return x.childrenI.len

proc `[]`*(x: XlNode, i: BackwardsIndex): XlNode {.inline.} =
  x[x.count - int(i)]

proc findPrefix(x: XlNode, ns: string): string =
  assert x.kind == xnElement
  var x {.cursor.} = x
  while x != nil:
    if x.nsmap != nil:
      for prefix, n in x.nsmap:
        if ns == n:
          return prefix
    x = x.parent

proc findNs(x: XlNode, prefix: string): string =
  assert x.kind == xnElement
  var x {.cursor.} = x
  while x != nil:
    if x.nsmap != nil:
      for p, ns in x.nsmap:
        if p == prefix:
          return ns
    x = x.parent

proc nsnameToRawName*(x: XlNode, tag: XlNsName): string =
  var prefix = ""
  if x != nil:
    prefix = x.findPrefix(tag.ns)

  if prefix == "":
    return tag.name

  else:
    return prefix & ':' & tag.name

proc tag*(x: XlNode): XlNsName =
  # return full namespace tag of a node
  assert x.kind == xnElement
  if x.nstag.name != "":
    return x.nstag

  # tag start with prefix
  let sp = x.fTag.split(':', maxsplit = 1)
  if sp.len == 2:
    let ns = findNs(x, sp[0])
    if ns != "":
      x.nstag = XlNsName(ns: ns, name: sp[1])
      return x.nstag

  # find default namespace
  let ns = findNs(x, "")
  if ns != "":
    x.nstag = XlNsName(ns: ns, name: x.fTag)
    return x.nstag

  # no namespace
  x.nstag = XlNsName(ns: "", name: x.fTag)
  return x.nstag

proc child*(x: XlNode, tag: XlNsName): XlNode =
  # find first child match tag
  for child in x.children:
    if child.tag == tag:
      return child

proc find*(x: Xlnode, tag: XlNsName): int =
  var index = 0
  for child in x.children:
    if child.tag == tag:
      return index
    index.inc
  return -1

iterator find*(x: XlNode, tag: XlNsName): XlNode =
  for child in x.children:
    if child.tag == tag:
      yield child

proc contains*(x: XlNode, tag: XlNsName): bool {.inline.} =
  return x.find(tag) >= 0

proc add*(father, son: XlNode) =
  # father must be xnElement
  # son can be other (xnText, etc)
  father.prepareChildren()
  add(father.s, son)
  if son.kind == xnElement:
    add(father.childrenI, father.s.len-1)
    son.parent = father

proc insert*(father, son: XlNode, index = -1) =
  # son must be xnElement
  assert son.kind == xnElement
  father.prepareChildren()
  if index < 0 or index >= father.childrenI.len:
    add(father, son)

  else:
    insert(father.s, son, father.childrenI[index])
    son.parent = father
    father.childrenOk = false

proc replace*(father, son: XlNode, index = -1) =
  var found = father.find(son.tag) # <- check xnElement both father and son
  if found < 0:
    father.insert(son, index)

  else:
    let index = father.childrenI[found]
    father.s[index].parent = nil

    son.parent = father
    father.s[index] = son

proc delete*(x: XlNode, i: Natural) =
  x.prepareChildren()
  let index = x.childrenI[i]
  x.s[index].parent = nil
  x.s.delete(index)
  x.childrenOk = false

proc delete*(x: XlNode, tag: XlNsName) =
  let index = x.find(tag)
  if index >= 0:
    x.delete(index)

proc newXlRootNode*(tag: XlNsName, prefix = ""): XlNode =
  if prefix == "":
    result = newElement(tag.name)
  else:
    result = newElement(prefix & ':' & tag.name)
  result.addNameSpace(tag.ns, prefix)

proc newXlNode*(x: XlNode, tag: XlNsName, kvs: varargs[tuple[key, val: string]]): XlNode =
  # return a new node with ns and name as tag
  # need a parent node (ns) to decide the raw tag name
  assert x.kind == xnElement
  result = newElement(nsnameToRawName(x, tag))
  result.parent = x # this is necessary so that we can get the nstag before add into parent
  if kvs.len != 0:
    result.attrs = kvs.toXlAttributes

proc newXlNode*(x: XlNode, tag: XlNsName, text: string): XlNode =
  # return a new text node
  result = newXlNode(x, tag)
  result.s.add newText(text)

proc addChildXlNode*(x: XlNode, tag: XlNsName, index = -1,
    kvs: varargs[tuple[key, val: string]]): XlNode {.inline.} =
  result = newXlNode(x, tag, kvs)
  insert(x, result, index)

proc addChildXlNode*(x: XlNode, tag: XlNsName, text: string,
    index = -1): XlNode {.inline.} =
  result = newXlNode(x, tag, text)
  insert(x, result, index)

template hasChild*(x: XlNode, tag: XlNsName, sym: untyped): bool =
  var sym {.cursor.}: XlNode
  if x.isNil:
    false
  else:
    sym = x.child(tag)
    sym != nil

template hasChild*(x: XlNode, tag: XlNsName): bool =
  if x.isNil:
    false
  else:
    x.child(tag) != nil

template hasChildN*(x: XlNode, tag: XlNsName, n: int, sym: untyped): bool =
  # provide symbol = child(tag)[n] (grandson)
  var sym {.cursor.}: XlNode
  if x.isNil:
    false
  else:
    let child {.cursor.} = x.child(tag)
    if child != nil and n < child.count:
      sym = child[n]
      true
    else:
      false

template editChild*(x: XlNode, tag: XlNsName, sym: untyped, index: int, body: untyped): untyped =
  # provide new child if not exists
  var sym {.cursor.} = x.child(tag)
  if sym.isNil:
    sym = addChildXlNode(x, tag, index)

  body

proc `[]`*(x: XlNode, key: XlNsName): string =
  if x.attrs.isNil:
    return ""

  let defaultNs = findNs(x, "")

  for attr, val in x.attrs:
    if defaultNs != "" and key == XlNsName(ns: defaultNs, name: attr):
      return val

    let sp = attr.split(':', maxsplit = 1)
    if sp.len == 2:
      let ns = findNs(x, sp[0])
      if ns != "" and key == XlNsName(ns: ns, name: sp[1]):
        return val

  # attribute without namespace can match any namespace?
  return x[key.name]

proc `[]=`*(x: XlNode, key: XlNsName, value: string) {.inline.} =
  x[nsnameToRawName(x, key)] = value

proc deleteAttr*(x: XlNode, name: string) {.inline.} =
  if x != nil and x.attrs != nil:
    x.attrs.del name

template hasAttr*(x: XlNode, attr: XlNsName|string, sym: untyped): bool =
  var sym = x[attr]
  sym != ""

template hasIntAttr*(x: XlNode, attr: XlNsName|string, sym: untyped): bool =
  var sym: int
  try:
    sym = parseInt(x[attr])
    true
  except:
    false

template hasFloatAttr*(x: XlNode, attr: XlNsName|string, sym: untyped): bool =
  var sym: float
  try:
    sym = parseFloat(x[attr])
    true
  except:
    false

template hasTrueAttr*(x: XlNode, attr: XlNsName|string): bool =
  try:
    let val = x[attr]
    val == "true" or (bool parseInt(val))
  except:
    false

proc dup*(x: XlNode): XlNode =
  case x.kind
  of xnElement:
    result = newElement(x.fTag)
    result.nstag = x.nstag # can copy

    if x.attrs != nil:
      for key, val in x.attrs:
        result[key] = val

    if x.nsmap != nil:
      for k, v in x.nsmap:
        result{k} = v

    for i in x.s:
      let child = i.dup()
      if child.kind == xnElement:
        child.parent = result
      result.s.add child

  of xnText: result = newText(x.fText)
  of xnCData: result = newCData(x.fText)
  of xnEntity: result = newEntity(x.fText)

template hasRawChild*(x: XlNode, tag: string, sym: untyped): bool =
  var sym {.cursor.} : XlNode
  sym = x.rawChild(tag)
  sym != nil

proc find*(x: var XmlParser, kind: XmlEventKind): bool =
  while true:
    if x.kind == kind: return true
    elif x.kind in {xmlError, xmlEof}: return false
    x.next()

proc find*(x: var XmlParser, tags: openArray[string]): int =
  while true:
    case x.kind
    of xmlElementStart, xmlElementOpen:
      for i, tag in tags:
        if x.elementName == tag:
          return i

    of xmlError, xmlEof:
      return -1

    else:
      discard

    x.next()

proc find*(x: var XmlParser, tag: string): bool =
  while true:
    case x.kind
    of xmlElementStart, xmlElementOpen:
      if x.elementName == tag:
        return true

    of xmlError, xmlEof:
      return false

    else:
      discard

    x.next()

proc parseIgnoreChild*(x: var XmlParser): XlNode =
  assert x.kind in {xmlElementStart, xmlElementOpen}
  result = newElement(x.elementName)
  while true:
    x.next()
    if x.kind == xmlAttribute:
      if x.attrKey == "xmlns":
        result{""} = x.attrValue
      elif x.attrKey.startsWith("xmlns:"):
        result{x.attrKey[6..^1]} = x.attrValue
      result[x.attrKey] = x.attrValue
    else:
      break
