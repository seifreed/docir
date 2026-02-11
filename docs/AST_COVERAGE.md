# AST Coverage Matrix

This document defines the working definition of "AST completo" for each supported format.
The goal is deterministic, semantic IR coverage of all meaningful parts, with security-relevant
content modeled explicitly.

## Legend
- Status: DONE, PARTIAL, TODO
- Coverage: structural coverage in IR (not 1:1 XML)
- Tests: concrete fixtures and validation rules

## OOXML (DOCX)
- Parts
  - word/document.xml -> Document/Section/Paragraph/Run/Table/Hyperlink (DONE)
  - word/styles.xml -> StyleSet (DONE)
  - word/numbering.xml -> NumberingSet (DONE)
  - word/headers*.xml -> Header (DONE)
  - word/footers*.xml -> Footer (DONE)
  - word/comments.xml -> Comment + ranges (DONE)
  - word/footnotes.xml, word/endnotes.xml -> Footnote/Endnote (DONE)
  - word/settings.xml, webSettings.xml, fontTable.xml -> WordSettings/WebSettings/FontTable (DONE)
  - drawingML/vml -> Shape/VmlShape/DrawingPart (DONE)
  - charts/smartart -> ChartData/SmartArtPart (DONE)
  - macros/ActiveX/OLE -> MacroProject/ActiveXControl/OleObject (DONE)
- Tests
  - fixtures/OOXML: full round-trip coverage + security scan (TODO)

## OOXML (XLSX)
- Parts
  - xl/workbook.xml -> Document/Worksheet (DONE)
  - xl/worksheets/*.xml -> Worksheet/Cell/Table (DONE)
  - xl/sharedStrings.xml -> SharedStringTable (DONE)
  - xl/styles.xml -> SpreadsheetStyles (DONE)
  - xl/pivotTables/*.xml -> PivotTable/PivotCache (DONE)
  - xl/calcChain.xml -> CalcChain (DONE)
  - external links, connections, query tables -> ExternalLinkPart/ConnectionPart/QueryTablePart (DONE)
  - macros/XLM -> MacroProject/XlmMacro (DONE)
- Tests
  - fixtures/OOXML: calc, pivot, formulas, macros (TODO)

## OOXML (PPTX)
- Parts
  - ppt/presentation.xml -> PresentationInfo/Slide (DONE)
  - ppt/slides/*.xml -> Slide/Shape (DONE)
  - ppt/slideLayouts/*.xml -> SlideLayout (DONE)
  - ppt/slideMasters/*.xml -> SlideMaster (DONE)
  - media, animations, transitions -> MediaAsset/SlideAnimation/SlideTransition (DONE)
  - comments/people/tags -> PptxComment/PptxCommentAuthor/PresentationTag (DONE)
- Tests
  - fixtures/OOXML: media + animations + OLE (TODO)

## ODF (ODT/ODS/ODP)
- Parts
  - content.xml -> Document/Section/Paragraph/Run/Table/Cell/Slide/Shape (DONE)
  - styles.xml -> StyleSet (DONE)
  - meta.xml -> DocumentMetadata (DONE)
  - settings.xml -> ExtensionPart + diagnostics (DONE)
  - scripts/objects -> MacroProject/OleObject (DONE)
- Tests
  - fixtures/ODF: text, calc, draw, external links (TODO)

## HWP (legacy)
- Parts
  - FileHeader/DocInfo/BodyText -> Section/Paragraph/Run (PARTIAL)
  - BinData -> MediaAsset (DONE)
  - Scripts -> MacroProject (PARTIAL)
- Tests
  - fixtures/HWP: encrypted + content streams (TODO)

## HWPX
- Parts
  - Contents/section*.xml -> Section/Paragraph/Run/Table (DONE)
  - Contents/header*.xml, footer*.xml -> Header/Footer (DONE)
  - Contents/masterPage*.xml -> Section (DONE)
  - Contents/content.hpf -> StyleSet (PARTIAL)
  - BinData -> MediaAsset (DONE)
  - External links/OLE -> ExternalReference/OleObject (PARTIAL)
- Tests
  - fixtures/HWPX: headers, tables, styles, objects (TODO)

## RTF
- Parts
  - Plain text -> Paragraph/Run (DONE)
  - Tables -> Table/TableRow/TableCell (PARTIAL)
  - Lists -> NumberingInfo (DONE)
  - Stylesheet -> StyleSet (DONE)
  - Fields -> Field/Hyperlink (DONE)
  - Objects/OLE -> OleObject (DONE)
  - Images -> MediaAsset (DONE)
  - Borders/shading -> TableCellProperties (PARTIAL)
- Tests
  - fixtures/RTF: lists + styles + tables + objects (PARTIAL)

## Validation rules
- Deterministic ordering in IR store
- Coverage: no unclassified parts for OOXML/ODF/HWPX
- Security nodes: external refs, OLE, macros explicitly mapped

