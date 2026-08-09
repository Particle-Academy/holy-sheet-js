# Changelog

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
