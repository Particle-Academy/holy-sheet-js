# Changelog

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
