# Changelog

All notable changes to `@particle-academy/holy-sheet-js` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.4.3] — 2026-09-14

### Fixed

- **A `columnWidths` key that is not a column index wrote a `NaN` column.** The
  normalizer read keys with `parseInt`, so `{"abc": 999}` became a width for
  column `NaN` in the written file. PHP overwrote column A with it and Python
  raised; all three follow one rule now, mirroring PHP holy-sheet 2.3.4
  (`src/schema/column-widths.ts`):
  - a **key** is a 0-based column index from 0 to 16383, as a string of digits;
  - a **width** is a non-negative finite number, or a string of digits (`"80.5"`).
  - `validate()` reports each other entry by path (`sheets[0].columnWidths.abc`),
    so `write()` and `toBytes()` refuse it.
  - `validateAndRepair()` turns a one- or two-letter key into its index (`"B"` is
    1) and drops an entry it cannot repair, and lists both.
  - The normalizer skips such an entry.

  **What you must do:** nothing, unless a schema carried a width keyed by a letter
  or by junk. `write()` now refuses it; run it through `validateAndRepair()`, or
  key widths by index.

## [2.4.2] — 2026-09-14

### Fixed

Six defects in the op code, mirroring `particle-academy/holy-sheet` 2.3.2 (five,
found by the Python port) and 2.3.3 (one), which are the reference. Each has a
test ported from PHP's `SheetOpsTest.php`, and each fails against 2.4.1 except
the second, which this port never had.

- **`Agent.opSchema()` rejected a `set_column_widths` op carrying a list.** 2.4.1
  allowed only an empty array, but PHP encodes widths keyed 0..n-1 as a list
  (`[120, 80, 140]`), and a schema read from PHP's JSON carries that list through
  this port's `diff()` into its op. `columnWidths` is now an object or an array
  of non-negative numbers, a list indexed by position. The schema matches PHP
  2.3.2's byte for byte, description included.
- **A `type` that is not an op type is skipped before anything else.** PHP's
  loose `switch` let `type: true` remove a sheet. This port's `switch` was
  always strict, so it never had that defect; `reduce()` now checks the type
  against `SheetOpSchema.TYPES` first, as PHP does.
- **A padded address wrote a key of its own.** `set_cell` stored `" a1 "` under
  `" A1 "`, and `clear_cell` could not reach A1 with it. Both now trim with PHP's
  `trim()` set (space, tab, LF, CR, NUL, vertical tab; not NBSP) before
  upper-casing.
- **A column-width key that is not a column index was read as column A** by
  `insert_columns` and `delete_columns` (`"abc"` became column 0), and could
  overwrite column A's width. It is dropped. A digit string (`"007"`) and a
  negative integer key are still indexes, as they are in PHP.
- **Two values JSON cannot hold compared as the same.** `SheetDiff.same()`, and
  so `diff()`, wrote `NaN`, `Infinity`, functions and symbols as `null`, so a
  cell going from `NaN` to `Infinity` recorded no change. It now throws a
  `TypeError` for a non-finite number, a lone UTF-16 surrogate in a string or a
  key (a JS string cannot hold invalid UTF-8, and this is its equivalent), a
  function or a symbol. Nesting deeper than PHP's `json_encode` depth of 4096
  throws too, at the same level PHP does (every object and array counts, an
  empty one included). The comparison is now a loop: the recursion it replaces
  overflowed the call stack near that depth, at a level that varied from run to
  run.
- **An op with a position or count that is not a number moved or unfroze
  things** (PHP 2.3.3). `add_sheet.index`, `move_sheet.toIndex`,
  `set_frozen.rows`/`cols` and the row and column ops' `at`/`count` were cast as
  PHP's `(int)` does, so `toIndex: "last"` moved a sheet to the front,
  `index: "end"` inserted one there, and `rows: "one"` unfroze the panes. A
  present value that is not an integer or a string of digits, `null` included,
  now skips the op; absent keys keep their defaults (`add_sheet` appends,
  `move_sheet` stays, `set_frozen` uses 0). As in PHP, the check covers every op
  that reaches a sheet, so a junk `count` on a `set_cell` skips it too. One
  divergence cannot be removed: `JSON.parse` gives `2` for `2.0`, so a position
  written `2.0` applies here, where PHP decodes a float and skips the op.

`tests/sheet-ops-parity.test.ts` gains PHP-checked reduce cases for each fix,
diff cases for a list of widths, a padded cell key and a width key that is not
an index, and `SheetDiff.same` at 4095 to 4097 levels, which the PHP helper
builds itself so the nesting never crosses JSON.

**What you must do:** nothing, unless you relied on one of the above. `diff()`
output changes only for a schema holding a padded cell key or a width key that
is not an index, where it matches PHP 2.3.3's, and `diff()` now throws on a
schema holding a value JSON cannot, which JSON input never does.

## [2.4.1] — 2026-09-14

### Fixed

- **`Agent.opSchema()` accepts a `set_column_widths` op with an empty array**, as
  PHP 2.3.1 now declares. PHP's diff emits `columnWidths: []` when every width is
  removed (its JSON for an empty map), which the schema rejected; an op stream
  written by the PHP package failed validation here. `columnWidths` may be an
  object or an empty array. The parity suite compares the schema with PHP's
  byte for byte, so the two stay one schema.

  **What you must do:** nothing.

## [2.4.0] — 2026-09-14

### Added

- **`Agent.diff()`, `Agent.reduce()`, `Agent.opSchema()` and `Agent.equivalent()`:
  a workbook's versions stored as ops** (holy-sheet [#7](https://github.com/Particle-Academy/holy-sheet/issues/7)).
  Ported from `particle-academy/holy-sheet` 2.3.0, which is the reference: the
  same algorithm, so the same two schemas give the same ops in the same order
  in both runtimes, and an op history written by one replays in the other.
  Hashing xlsx bytes cannot keep a one-cell edit small, because a zip changes
  nearly every byte, so a version history had to store a whole file per edit.
  `diff(newer, older)` is the op list that restores `older` from `newer`.
  - `reduce(a, diff(a, b))` equals `b`, key order aside. The ops are verified
    by replaying them: a sheet the granular ops cannot reproduce is replaced
    whole, and so, as a last resort, is the workbook.
  - One changed cell is one `set_cell`. Rows and columns are aligned by content
    first, so an inserted row is one `insert_rows` plus its cells rather than
    every cell below it rewritten.
  - Schemas that write the same workbook diff to `[]`, so a save without a change
    records nothing. A columns/rows sheet and the cells it becomes are the same,
    and so is the creation time the writer stamps on a schema that names none.
  - `set_cell`, `set_range` and `set_workbook` are fancy-sheets' `SheetOp` shapes
    and behave as its reducer does (a `set_cell` without a formula clears it and
    keeps the format). The rest are holy-sheet's: `clear_cell`,
    `insert_rows`/`delete_rows`, `insert_columns`/`delete_columns`,
    `add_sheet`/`remove_sheet`/`rename_sheet`/`move_sheet`/`replace_sheet`,
    `set_merged_regions`, `set_column_widths`, `set_frozen` and `set_meta`.
  - Row and column ops move cells, merged regions and column widths. They do not
    rewrite formula text; a formula that changes with an insert is its own
    `set_cell`.
- **`SheetOp`**, a discriminated union on `type` (and one interface per op),
  plus **`SheetDiff`**, **`SheetReducer`** and **`SheetOpSchema`**, exported
  beside the other building blocks.

`tests/sheet-ops-parity.test.ts` runs the PHP package's `Agent::diff` and
`Agent::reduce` on every case and requires the same ops, field for field and in
order: each edit PHP's own suite pins, both ways; seeded random cell edits;
seeded row and column inserts and deletes; moved rows and raw alignments, where
the delete-first tie-break decides the result; and the op JSON Schema, byte for
byte. Two differences cannot be removed, because JS values do not carry the
distinction:

- PHP counts `1` and `1.0` as different values, so its diff rewrites a `1` as
  `1.0` with a `set_cell` where this one records nothing.
- An emptied column-width map is emitted as `{}`, where PHP's JSON writes its
  empty array as `[]`. Both reducers read either as empty.

**What you must do:** nothing. This only adds methods.

## [2.3.1] — 2026-09-14

### Fixed

- **`Agent.lint()` accepts quoted sheet names** (holy-sheet [#6](https://github.com/Particle-Academy/holy-sheet/issues/6)).
  `=SUM('My Earnings Projection'!A2:A3)` linted as `#NAME?` because the
  tokenizer had no case for `'`, so any cross-sheet formula pointing at a sheet
  whose name contains a space failed. Excel's doubled-quote escape works too:
  `'Q3 ''Final'''!B2` is the sheet `Q3 'Final'`.
- **A reference to a sheet that does not exist is `#REF!`**, quoted or not, and
  the hint names the sheet and lists the ones that exist. It used to lint clean,
  because the missing sheet's cells read as blanks.
- **Sheet names match case-insensitively**, as in Excel.

The same fix as `particle-academy/holy-sheet` 2.2.1, with the hint text
identical in all three runtimes and pinned by a test in each.

**What you must do:** nothing, unless you relied on a formula that names a
missing sheet linting clean. It now reports `#REF!`.

## [2.3.0] — 2026-09-13

### Added

- **`Agent.read()` and `Agent.describe()` read OpenDocument spreadsheets
  (`.ods`) into the same schema as `.xlsx`.** Ported from
  `particle-academy/holy-sheet`'s new reader, which has the full list of what
  maps (values, formulas translated to A1, repeats, merges, comments, styles,
  data styles, metadata) and what does not (frozen panes, column widths,
  fonts, function-name translation). The format is sniffed from the bytes, so
  the caller's branch on the file type can go.

  Diffed against the PHP reader on the PHP repo's own fixtures, key order
  included, in `tests/ods-reader-parity.test.ts`: a LibreOffice-converted
  workbook, a hand-written native file, and a hand-built package of what
  LibreOffice rewrites on save. `tests/ported-php/OdsReaderTest.test.ts` ports
  the PHP assertions. Both load the fixtures from the PHP checkout
  (`HOLY_SHEET_PHP_SRC`, or the sibling directory), as the parity suites
  already did.

- **`UnsupportedFormatException`**, exported, thrown for bytes that are
  neither format, with the declared `mimetype` when there is one.
- **`OdsReader`** and **`FormatSniffer`**, exported beside `XlsxReader`, and
  `XlsxReader.readFiles()` for an already-unzipped package.
- **`parseXml(src, { namespaces: true })`** resolves namespaces and keeps mixed
  content in order. OpenDocument needs both; without the option the parser
  behaves exactly as before.

### Changed

- **Unreadable bytes now throw `UnsupportedFormatException` instead of a bare
  `Error`.** It extends `Error`, so **an existing `catch` keeps working: do
  nothing.** The message for bytes that are not a zip still says "not a zip
  archive"; only code matching the rest of the old text (`(no EOCD record)`,
  `missing xl/workbook.xml`) sees different wording.

### Fixed

- **The tool schema announced its own shipped features as unreleased.**
  `src/holy-sheet.schema.json` is a byte-identical copy of the PHP twin's, and
  carried the same defect: every formatting field described in the future tense
  ("Lands in 0.3.", "Lands in 0.5.") under a heading declaring a v0.2.0 that
  could not style anything, while the package shipped at 2.x with all of it
  working.

  Updated in lockstep with `particle-academy/holy-sheet`; both checksum pins
  moved together. See that package's entry for the measurement.

- **An empty `cells: []` validates.** PHP's `describe()` reports a sheet with no cells as `cells: []`, and this validator rejected any array, so a schema read in PHP and written here failed on an empty sheet. The PHP validator had the same defect and is fixed in holy-sheet 2.2.0. A non-empty list is still an error.

## 2.2.1 — 2026-09-10

### Fixed

- **`version()` reports the version this package actually ships as.** It
  returned `1.0.0` from a 2.2.x release. The constant had drifted because
  nothing compared it to the packaging metadata — the same shape as every other
  two-copies-of-one-number failure in this estate.

  `VersionIsSingleSourcedTest` / `version.test.ts` now pins it, so the class is
  closed rather than the instance fixed. `dark-slide-py` already had that
  assertion and was the only engine in the family to catch itself.


## 2.2.0 — 2026-09-10

### Added

- **A checksum test pinning `holy-sheet.schema.json` to its PHP twin.** The
  schema is a byte-identical copy of `holy-sheet/skills/holy-sheet.schema.json`,
  kept in sync by remembering to edit both. Nothing checked that — and this is
  the file handed to an LLM as the tool definition, so a one-sided edit did not
  fail a build, it changed what an agent was told the API is on one backend and
  not the other.

  To change the schema: edit both copies, run either suite, paste the new hash
  into both tests. The mild annoyance is the mechanism.

  The hash is taken over **newline-normalised** content. Neither repo has a
  `.gitattributes`, so the file is stored LF and lands CRLF on a Windows
  checkout — a raw-bytes checksum would have passed on CI and failed on the
  maintainer's own machine, which is a worse failure than the one being
  prevented.

### Fixed

- **A sheet with BOTH a table and explicit cells no longer discards the table.**
  `normalizeSheet` returned the moment `cells` was set, silently throwing away
  `columns`, `rows`, `totals` and `theme`. A four-row report with one styled
  title cell wrote a **one-cell workbook**, and `validate()` reported no errors,
  because the validator explicitly permits both shapes together.

  The PHP twin had the identical bug in the identical place. That is why no
  parity test caught it: **parity detects disagreement, and the two runtimes
  agreed.**

  **What you must do: nothing.** A sheet using only one of the two shapes
  behaves exactly as before. Explicit cells win at their address, and their
  format is merged over the table's, so a cell asking only for `bold` keeps the
  column's currency format instead of dropping to bare.

- **Bare scalar cells are written instead of silently emptied.** `{A1: 42}` read
  `.value` off a number, got `undefined`, and wrote a blank cell — no error,
  just a hole where the number should be. **This one was Node-only:** PHP has
  always had the branch, and the port dropped it.

- **A bare `"=SUM(...)"` string is promoted to a real formula**, in both cell
  maps and rows. Also Node-only: the same schema produced a working formula in
  PHP and the literal text `=SUM(...)` in Node.

- **A `comment` given as a plain string is no longer dropped** — only the object
  form was read, and the string form was accepted by the validator and lost.

## 2.1.0 — 2026-08-09

### Fixed

- **`1e300` was written to the sheet as `1e+3`.** The worst bug in this package's
  history, and it never threw. `toFixed` switches to exponential notation at
  1e21, and the trailing-zero strip then chewed the EXPONENT rather than a
  fraction:

  ```
  (1e300).toFixed(14)         === "1e+300"
  "1e+300".replace(/0+$/, "") === "1e+3"
  ```

  So a cell holding 1e300 was written holding 1000. Not a crash, not invalid
  XML — a different number, in a file someone opens later and believes.

- **`1e21` and above were written as `<v>1e+21</v>`**, which is not valid cell
  content. This had a second cause the first fix did not reach: the writer only
  called `formatFloat` for NON-integers, and `Number.isInteger(1e21)` is `true`,
  so it took a `String(v)` branch. Mirroring PHP's `is_float(...)` test leaks —
  a PHP int never stringifies exponentially, an integral JS double does.

- **Numeric strings lost their exponent.** `numericStringToNumber` used
  `parseInt` when the string had no `.`, and `parseInt` stops at the `e`:

  | input | was | now |
  |---|---|---|
  | `"1e5"` | `1` | `100000` |
  | `"1E5"` | `1` | `100000` |
  | `"2e-3"` | `2` | `0.002` |
  | `"1e21"` | `1` | `1e21` |

  It is now `Number(v)` — which is what `FormulaLinter` already used for the
  same job, so the two disagreed with each other as well as with PHP.

- **`NaN` and `Infinity` were written into cells.** Both now write `0`, matching
  the PHP twin.

### Added

- A `numericHazards` case in the cross-engine parity fixtures. The unit tables on
  both sides were green while the engines still disagreed on `1e21`; only the
  byte-for-byte diff caught it. Numbers are now a parity *case*, not a
  convention.

### Changed

- **BREAKING in effect, though the API is identical:** values listed above now
  serialize differently, because they were wrong. If you have a golden xlsx
  containing an exponent-notation string, a value ≥ 1e21, or `NaN`, regenerate
  it. Anything below 1e21 with no exponent string is byte-identical to 2.0.2.

## 2.0.2 — 2026-08-09

### Fixed

- **`scripts/php-describe.php` was missed by 2.0.1**, so the reader-parity suite
  still could not find the PHP classes outside the `.agi` envelope and CI failed
  on the 2.0.1 commit. It now honours `HOLY_SHEET_PHP_SRC` like its sibling.

  Worth saying plainly: CI caught this the very first time the parity suites
  actually ran. Under the old `skipIf` the incomplete fix would have gone green.

## 2.0.1 — 2026-08-09

### Fixed

- **The cross-engine parity suites now actually run in CI.** They are the
  strongest guarantee this pair has — every OOXML part diffed byte-for-byte
  against the PHP `holy-sheet` — and they had never executed on a runner.

  Two things combined to hide it. `describe.skipIf(!HAS_PHP)` skipped silently,
  and `ci.yml` installed Node only, so the build went green with **zero
  cross-engine coverage**. Separately, the parity helper autoloaded the PHP
  sources from a hard-coded `../../holy-sheet/src`, which only resolves inside
  the `.agi` envelope — so even with PHP present, a different layout found no
  classes.

  Now: CI installs PHP 8.4 and checks out the PHP repo; the helper takes
  `HOLY_SHEET_PHP_SRC` (falling back to the sibling path); and a missing PHP
  **throws in CI** instead of skipping. A green build that asserts nothing is
  worse than a red one, because nobody investigates green.

  **What you must do:** nothing, unless you run these suites outside the
  envelope — then set `HOLY_SHEET_PHP_SRC` to the PHP package's `src`.

## 2.0.0 — 2026-08-07

### Changed

- **BREAKING — Node 18 is no longer supported.** `engines.node` moves from `>=18` to `>=22`.

  **What you must do:** on Node 22 or newer, nothing. Note npm only *warns* on an `engines` mismatch while **pnpm fails the install**, so this surfaces differently depending on your package manager. Node 18 is end-of-life and 20 is maintenance-only.

### Why

These are the kit 0.5 platform floors, applied across every package at once so a consumer never has to resolve a mix. **No API changed, nothing was removed, nothing was renamed** — only what the package requires.

This package is past 1.0, so a floor raise takes a **major**. Most of the suite is pre-1.0 and lands the identical change in a minor — that is semver, not a difference in how much changed.

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
