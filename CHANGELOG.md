# Changelog

All notable changes to `@particle-academy/holy-sheet-js` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
