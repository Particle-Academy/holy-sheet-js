/**
 * An OpenDocument `table:formula` -> the A1 syntax the xlsx reader returns.
 * Mirrors PHP `Reader\Ods\OdsFormula`; the full list of what is and is not
 * translated is documented there.
 *
 * `of:=SUM([.A1:.B2];[$'Other Sheet'.C3])` becomes `SUM(A1:B2,'Other Sheet'!C3)`.
 */
export function odsFormulaToA1(formula: string): string {
  let grammar = "of";
  const m = /^(of|oooc|msoxl):/i.exec(formula);
  if (m) {
    grammar = m[1]!.toLowerCase();
    formula = formula.slice(m[0].length);
  }
  if (formula.startsWith("=")) formula = formula.slice(1);

  return grammar === "msoxl" ? formula : translate(formula);
}

function translate(s: string): string {
  let out = "";
  const length = s.length;
  let arrayDepth = 0;

  for (let i = 0; i < length; i++) {
    let c = s[i]!;

    if (c === '"') {
      let end = i + 1;
      while (end < length) {
        if (s[end] === '"') {
          if (end + 1 < length && s[end + 1] === '"') {
            end += 2;
            continue;
          }
          break;
        }
        end++;
      }
      out += s.slice(i, end + 1);
      i = end;
      continue;
    }

    if (c === "[") {
      let end = i + 1;
      let quoted = false;
      while (end < length && (quoted || s[end] !== "]")) {
        if (s[end] === "'") quoted = !quoted;
        end++;
      }
      out += reference(s.slice(i + 1, end));
      i = end;
      continue;
    }

    if (c === "{") {
      arrayDepth++;
    } else if (c === "}") {
      arrayDepth = Math.max(0, arrayDepth - 1);
    } else if (c === ";") {
      c = ",";
    } else if (c === "|" && arrayDepth > 0) {
      c = ";";
    }
    out += c;
  }

  return out;
}

/** The inside of one `[...]`. */
function reference(ref: string): string {
  const parts: string[] = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i < ref.length; i++) {
    if (ref[i] === "'") {
      quoted = !quoted;
    } else if (ref[i] === ":" && !quoted) {
      parts.push(ref.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(ref.slice(start));

  const parsed = parts.length <= 2 ? parts.map(part) : [null];
  if (parsed.some((p) => p === null)) {
    return ref.includes("#REF!") ? "#REF!" : `[${ref}]`;
  }

  const [sheet, address] = parsed[0]!;
  if (parsed.length === 1) return sheetPrefix(sheet) + address;

  const [sheet2, address2] = parsed[1]!;
  if (sheet2 === null || sheet2 === sheet) return sheetPrefix(sheet) + address + ":" + address2;
  if (sheet === null) return address + ":" + sheetPrefix(sheet2) + address2;

  // A range across sheets: one prefix naming both, quoted as a whole.
  const both =
    needsQuotes(sheet) || needsQuotes(sheet2)
      ? "'" + sheet.replace(/'/g, "''") + ":" + sheet2.replace(/'/g, "''") + "'"
      : sheet + ":" + sheet2;

  return both + "!" + address + ":" + address2;
}

/** One end of a reference: `$'Sheet'.$A$1`, `Sheet.A1`, `.A1`, `.A`, `.1`. */
function part(p: string): [string | null, string] | null {
  const length = p.length;
  let i = 0;
  if (i < length && p[i] === "$") i++;

  let sheet: string | null = null;
  if (i < length && p[i] === "'") {
    sheet = "";
    i++;
    for (;;) {
      if (i >= length) return null;
      if (p[i] === "'") {
        if (i + 1 < length && p[i + 1] === "'") {
          sheet += "'";
          i += 2;
          continue;
        }
        i++;
        break;
      }
      sheet += p[i++];
    }
    if (i >= length || p[i] !== ".") return null;
    i++;
  } else {
    const dot = p.indexOf(".", i);
    if (dot < 0) return null;
    if (dot > i) sheet = p.slice(i, dot);
    i = dot + 1;
  }

  const address = p.slice(i);
  if (!/^(\$?[A-Za-z]{1,3}\$?[0-9]+|\$?[A-Za-z]{1,3}|\$?[0-9]+)$/.test(address)) return null;

  return [sheet, address];
}

function sheetPrefix(sheet: string | null): string {
  if (sheet === null) return "";
  return (needsQuotes(sheet) ? "'" + sheet.replace(/'/g, "''") + "'" : sheet) + "!";
}

/** Whether Excel needs the sheet name quoted. Over-quoting is still valid. */
function needsQuotes(sheet: string): boolean {
  return (
    !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ||
    /^[A-Za-z]{1,3}[0-9]+$/.test(sheet) ||
    /^[Rr]([0-9]+)?[Cc]([0-9]+)?$/.test(sheet)
  );
}
