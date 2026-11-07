# Changelog

## [1.0.2] - 2026-11-07

A security patch on the 0.4 kit line. It carries this fix and nothing else.

### Security

- **The reader now refuses a worksheet whose column-width range goes past
  Excel's column limit of 16,384.** A crafted file could declare a column-width
  range billions of columns wide, and reading it used memory until the process
  ran out. The reader now throws before it walks the range, with a message
  naming the limit. Real spreadsheets never go past column 16,384, so they read
  exactly as before.
  **What to do:** update. No code changes.

## 1.0.1 — 2026-11-04

A security patch on the 0.4 kit line. It carries this fix and nothing else.

### Security

- **The reader now refuses a `DOCTYPE` in every XML part of an xlsx, as the PHP
  engine does.** The parser skipped a declaration without reading it, which could
  let a crafted file change the values the reader returned. A part is also read
  in the encoding it declares, so a UTF-16 part is read correctly and is checked
  for what it really says.
  **What to do:** upgrade if you read files you did not write. No code changes.
  Real spreadsheets carry no `DOCTYPE`, so they read exactly as before. A part
  with one is treated as if it were malformed: a refused `xl/workbook.xml` still
  throws, and any other part is read as absent. So a refused worksheet gives an
  empty sheet with its name, and refused shared strings give blank string cells.

## 1.0.0 — 2026-05-30

Initial Node/TypeScript port of `particle-academy/holy-sheet` (PHP), at
**feature-parity with PHP 1.2.0**. Zero-dependency, isomorphic (browser + Node).

- Full `Agent` surface: `validate`, `toBytes`, `write` (Node), `validateAndRepair`,
  `lint`, `fromArray`, `fromCsv`, `read`/`describe`, `toolDefinition`, `version`,
  plus the `HolySheet` instance class.
- xlsx **writer** — inline-string/number/bool/date/formula cells, deduped
  styles (fonts/fills/borders/numFmts/xfs), currency/percent/date number
  formats, 4 themes, symbolic totals, merged cells, frozen panes, column
  widths, comments (+ VML), multi-sheet.
- xlsx **reader** — round-trips the above back to a schema (own + Excel-authored
  files), via a hand-rolled isomorphic ZIP (STORE write / inflate read) and a
  tiny XML parser.
- **Formula linter** — `#VALUE!/#REF!/#DIV/0!/#NAME?/#CIRC!` with off-by-one
  hints; 15+ functions.
- **Verified byte-identical** to the PHP engine across sparse, themed/totaled,
  decorated, and multi-sheet workbooks (cross-engine parity suite).
