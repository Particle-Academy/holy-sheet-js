# AGENTS.md — holy-sheet (Node / TypeScript)

This file describes **this repo's code**: its API, its invariants, and the traps
in it. Process rules (publishing, versioning, backports, support lifecycle) live
in the envelope's `RULES.md` and are deliberately not repeated here: a copy in a
repo is frozen on whichever branch you have checked out.

`CLAUDE.md` is a symlink to this file. Edit this one.

## What this is

A zero-dependency, isomorphic `.xlsx` writer, reader and formula linter,
published as **`@particle-academy/holy-sheet`** (never with a `-js` suffix). It
is a port of one contract implemented three times:

| | |
|---|---|
| `particle-academy/holy-sheet` (PHP) | the reference: semantics are decided there |
| `@particle-academy/holy-sheet` (this) | Node/TypeScript port, bytes in / bytes out |
| `fancy-holy-sheet` (`holy-sheet-py`) | Python port, which mirrors this repo's module layout |

**PHP is normative for behaviour.** Where this port and PHP disagree, this port
is wrong, and the fix is to mirror PHP rather than to pick a third answer.

## The API

`Agent` (in `src/agent.ts`) mirrors PHP's static `Agent`: `validate`,
`validateAndRepair`, `toBytes`, `write`, `read`, `describe`, `lint`,
`fromArray`, `fromCsv`, `diff`, `reduce`, `opSchema`, `equivalent`,
`toolDefinition`, `guide`, `convertTemplate`, `version`. `HolySheet` wraps it as
a class. Lower-level classes keep their PHP names (`Validator`, `Normalizer`,
`FormulaLinter`, `XlsxWriter`, `XlsxReader`, `OdsReader`, `CellAddress`, ...)
and are exported from `src/index.ts`.

- **Universal methods are synchronous and run in a browser.** Only the
  file-touching ones (`write`, `describe` from a path) are `async` and
  Node-only, because a browser has no synchronous filesystem. Keep `node:*`
  imports out of everything else.
- **`read()` takes bytes** (`Uint8Array` or `ArrayBuffer`), and sniffs `.ods`
  as well as `.xlsx` into the same schema.
- **The input schema is a plain object and must stay one.** The declarative
  one-shot schema is the product; `Validator` is the gate and `Repairer` fixes
  the loose input a typed constructor would reject.

Not ported, deliberately: PHP's `Laravel/`, `Toolkit/` and `Schema\Dumper`.

## Invariants

**`VERSION` in `src/agent.ts` must equal `package.json`'s `version`.**
`tests/version.test.ts` fails otherwise.

**`src/holy-sheet.schema.json` is one of three byte-identical copies** (PHP's
`skills/holy-sheet.schema.json`, and the Python port's). Each repo pins its
SHA-256 over CRLF-normalised content (`tests/schema-sync.test.ts` here). Edit
one copy, edit all three, update all three constants. It is the file handed to
an LLM as the tool definition.

**Byte parity with PHP is the definition of done for the writer.**
`tests/parity.test.ts` runs the PHP writer and asserts byte-identical OOXML
PARTS for every fixture. It skips locally when `php` is not on `PATH` and
FAILS in CI, because a skip there would be a green build with no cross-engine
coverage. So:

- **The XML is built as strings.** Attribute order, self-closing style and the
  absence of whitespace between elements are part of the contract.
- **Cells in a row are emitted in ascending column INDEX** (`A..Z, AA, ...`),
  as OOXML requires (`src/writer/xlsx-writer.ts` sorts on
  `CellAddress.index`; PHP uses `uksort` on `CellAddress::index`). A bare
  `.sort()` is a string sort and writes `A, AA, AB, ..., B, C`, which is what
  all three engines did until 2026-10-28, and it lost data: Google Sheets
  silently drops every cell in a row after the first out-of-order one, so a
  wide holy-sheet row came back from a Sheets import/export holding only `A1`.
  Pinned by `tests/wide-row-cell-order.test.ts` and the `wide` parity fixture.
- **Never byte-compare the container.** This zip writer (`src/zip/`) pins a
  fixed 1980-01-01 DOS timestamp so its own output is reproducible; PHP writes
  through `ZipArchive` with real times. Only the parts can match.

**Numbers follow PHP, not JavaScript.** `src/util.ts` holds the PHP-semantics
helpers (`isNumericString`, numeric-string coercion). `parseInt` / `parseFloat`
do not mirror PHP's casts (`parseInt("1e5")` is 1, PHP says 100000), so reach
for the helper, never the builtin, anywhere a value is coerced.

**The ops (`src/ops/`) mirror PHP array semantics** (`src/ops/php.ts`): `diff`
must return PHP's op list in PHP's order for the same input, and PHP's answer
depends on array behaviour a JS object does not have.

## The formula linter (`src/schema/formula-linter.ts`)

A line-by-line port of PHP `Schema\FormulaLinter`. Read PHP's `AGENTS.md`
section on it for the design; the port-specific traps are these.

- **It evaluates while it parses**, and lazy `IF` depends on that: `parseIf`
  parses every argument but evaluates only the condition and the selected
  branch. The untaken branch runs with `skipping > 0`, in which every operator,
  `callFunction` and `resolveRefOrRange` returns `null`. A new operator or value
  source must honour `skipping`.
- **Error labels and hint texts are byte-identical to PHP and Python**, pinned
  by literal tests in `tests/ported-php/FormulaLinterTest.test.ts`. The hints
  contain em dashes and the doubled quote `"a""b"`.
- **Three JavaScript defaults had to be overridden to match PHP**, each pinned
  by a test:
  - character classes are ASCII (`isSpace`, `isDigit`, `isAlpha`), because
    `/\s/` accepts a no-break space that PHP's `ctype_space` rejects;
  - text compares by code point (`compareCodePoints`), because `<` on strings
    compares UTF-16 code units and puts an astral character below U+E000;
  - `ROUND` shifts the shortest decimal digits (`excelRound`), because
    `Math.round(x * 100)` rounds the binary double and gives `ROUND(1.005,2)`
    as 1.
- **`LEN` is `.length`**, which already counts UTF-16 code units, as Excel does.
  Python needed a fix for this; JavaScript did not.

## Testing

`npm test` (vitest) and `npm run lint` (`tsc --noEmit`). `tests/ported-php/`
holds tests ported case-for-case from the PHP suite, named after the PHP file.
