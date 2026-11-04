import { mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { Agent, unzipSync, zipSync } from "../src/index";
import { parseXml } from "../src/reader/xml";

/**
 * Every XML part of an uploaded file is read through one gate, which refuses a
 * DOCTYPE. Mirrors PHP tests/Unit/DoctypeRefusalTest.php.
 *
 * Everything goes through the public reader. This port's parser has no entity
 * machinery, so nothing can expand here: what the suite holds is that it refuses
 * the same parts as the other two, and reads no content they would not.
 */
const MARKER = "HSBASE";
const ENTITY = "HSINT";
const DOCTYPE = `<!DOCTYPE r [<!ENTITY x "${ENTITY}">]>`;

const enc = new TextEncoder();
const dec = new TextDecoder();
type Parts = Record<string, Uint8Array>;

function xlsxParts(): Parts {
  const parts = unzipSync(
    Agent.toBytes({
      meta: { creator: MARKER, created: "2026-01-01T00:00:00Z" },
      sheets: [
        {
          name: "Sheet1",
          cells: {
            A1: { value: MARKER, comment: { text: MARKER, author: "Auth" } },
            A2: { value: 1, format: { displayFormat: "percentage" } },
            // An object: this line's writer leaves a bare scalar in a `cells` map empty.
            A3: { value: 2 },
          },
        },
      ],
    } as never),
  );
  parts["xl/sharedStrings.xml"] = enc.encode(
    `<?xml version="1.0" encoding="UTF-8"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>${MARKER}</t></si></sst>`,
  );
  parts["xl/worksheets/sheet1.xml"] = enc.encode(
    dec.decode(parts["xl/worksheets/sheet1.xml"]!).replace(/<c r="A3"([^>]*)><v>2<\/v><\/c>/, '<c r="A3"$1 t="s"><v>0</v></c>'),
  );
  return parts;
}

function zipOf(parts: Parts): Uint8Array {
  return zipSync(Object.entries(parts).map(([name, data]) => ({ name, data })));
}

function withDoctype(xml: string, doctype: string = DOCTYPE): string {
  const m = /^<\?xml[^>]*\?>\s*/.exec(xml);
  return m ? m[0] + doctype + xml.slice(m[0].length) : doctype + xml;
}

const declared = (xml: string, encoding: string): string =>
  xml.replace(/^<\?xml[^>]*\?>/, `<?xml version="1.0" encoding="${encoding}"?>`);

const utf16le = (s: string): Uint8Array => new Uint8Array(Buffer.from(s, "utf16le"));
const utf16be = (s: string): Uint8Array => new Uint8Array(Buffer.from(s, "utf16le").swap16());
const concat = (a: number[], b: Uint8Array): Uint8Array => new Uint8Array([...a, ...b]);

const ENCODINGS: Record<string, (x: string) => Uint8Array> = {
  "plain UTF-8": (x) => enc.encode(x),
  "UTF-8 with a BOM": (x) => concat([0xef, 0xbb, 0xbf], enc.encode(x)),
  "UTF-16LE with a BOM": (x) => concat([0xff, 0xfe], utf16le(declared(x, "UTF-16"))),
  "UTF-16BE with a BOM": (x) => concat([0xfe, 0xff], utf16be(declared(x, "UTF-16"))),
  "UTF-16LE without a BOM": (x) => utf16le(declared(x, "UTF-16")),
  "past a long comment": (x) => enc.encode(x.replace(/^(<\?xml[^>]*\?>\s*)/, `$1<!-- ${"p".repeat(5000)} -->`)),
};
const ENCODING_NAMES = Object.keys(ENCODINGS);

/** The parts this port reads whose marker reaches the result. */
const MARKER_PARTS = ["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "docProps/core.xml", "xl/comments1.xml"];

const text = (parts: Parts, part: string): string => dec.decode(parts[part]!);
const referencing = (parts: Parts, part: string): string => text(parts, part).split(MARKER).join("&x;");

function readText(parts: Parts): string {
  try {
    return JSON.stringify(Agent.read(zipOf(parts)));
  } catch {
    return "(refused)";
  }
}

const count = (haystack: string, needle: string): number => haystack.split(needle).length - 1;

it("reads every marker from a clean file, so a missing one means refused", () => {
  // A1 value, comment, A3 shared string and creator. Styles has its own test:
  // this port returns no format code as text.
  expect(count(readText(xlsxParts()), MARKER)).toBe(4);
});

describe("a DOCTYPE on a part", () => {
  it.each(MARKER_PARTS.flatMap((p) => ENCODING_NAMES.map((e) => [p, e] as const)))("is refused in %s, %s", (part, encoding) => {
    const parts = xlsxParts();
    parts[part] = ENCODINGS[encoding]!(withDoctype(referencing(parts, part)));
    const got = readText(parts);

    expect(got).not.toContain(ENTITY);
    expect(got).not.toContain("&x;");
  });

  it.each(MARKER_PARTS.flatMap((p) => ENCODING_NAMES.map((e) => [p, e] as const)))("still reads %s without one, %s", (part, encoding) => {
    const parts = xlsxParts();
    const clean = readText(parts);
    parts[part] = ENCODINGS[encoding]!(text(parts, part));

    expect(readText(parts)).toBe(clean);
  });

  it("takes no number format from a refused styles part", () => {
    const parts = xlsxParts();
    expect(readText(parts)).toContain("percentage");
    parts["xl/styles.xml"] = enc.encode(withDoctype(text(parts, "xl/styles.xml")));

    expect(readText(parts)).not.toContain("percentage");
  });

  it.each([
    ["xl/_rels/workbook.xml.rels", "worksheets/sheet1.xml", '"Sheet1"'],
    ["xl/worksheets/_rels/sheet1.xml.rels", "../comments1.xml", '"comment"'],
  ])("does not follow %s", (part, target, probe) => {
    const parts = xlsxParts();
    expect(text(parts, part)).toContain(target);
    parts[part] = enc.encode(withDoctype(text(parts, part)));

    expect(readText(parts)).not.toContain(probe);
  });
});

describe("what a consumer sees, per part: the same table as PHP", () => {
  it("throws for a refused workbook part", () => {
    const parts = xlsxParts();
    parts["xl/workbook.xml"] = enc.encode(withDoctype(text(parts, "xl/workbook.xml")));

    expect(() => Agent.read(zipOf(parts))).toThrow("failed to parse xl/workbook.xml");
  });

  it("keeps the sheet, with no cells, for a refused worksheet part", () => {
    const parts = xlsxParts();
    parts["xl/worksheets/sheet1.xml"] = enc.encode(withDoctype(text(parts, "xl/worksheets/sheet1.xml")));
    const sheet = (Agent.read(zipOf(parts)) as any).sheets[0];

    expect(sheet.name).toBe("Sheet1");
    expect(sheet.cells ?? {}).toEqual({});
  });

  it("reads the other parts without refused shared strings", () => {
    const parts = xlsxParts();
    parts["xl/sharedStrings.xml"] = enc.encode(withDoctype(text(parts, "xl/sharedStrings.xml")));
    const cells = (Agent.read(zipOf(parts)) as any).sheets[0].cells;

    expect(cells.A1.value).toBe(MARKER);
    expect(cells.A3?.value ?? "").not.toBe(MARKER);
  });
});

describe("the subset is never read as content", () => {
  // An entity value may hold a '>', which is where a scan for the end of the
  // declaration would stop. What follows it must not become the part's root.
  const smuggling = (root: string): string => `<!DOCTYPE r [<!ENTITY y "a>${root}">]>`;

  it("takes no root from inside the subset", () => {
    expect(parseXml(`<?xml version="1.0"?>${smuggling("<evil/>")}<r/>`)).toBeNull();
  });

  it("takes no cells from inside the subset", () => {
    const parts = xlsxParts();
    const sheet = "<worksheet><sheetData><row r='1'><c r='A1' t='str'><v>HSSMUGGLED</v></c></row></sheetData></worksheet>";
    parts["xl/worksheets/sheet1.xml"] = enc.encode(withDoctype(text(parts, "xl/worksheets/sheet1.xml"), smuggling(sheet)));

    expect(readText(parts)).not.toContain("HSSMUGGLED");
  });
});

describe("external entities", () => {
  let dir: string;
  let url: string;

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), "hs-dt-"));
    const canary = join(dir, "canary.txt");
    writeFileSync(canary, "HSCANARY");
    url = `file:///${canary.replace(/\\/g, "/")}`;
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  it.each(MARKER_PARTS)("neither loads nor leaks a file entity, %s", (part) => {
    const parts = xlsxParts();
    parts[part] = enc.encode(withDoctype(referencing(parts, part), `<!DOCTYPE r [<!ENTITY x SYSTEM "${url}">]>`));
    const got = readText(parts);

    expect(readFileSync(url.replace("file:///", ""), "utf8")).toBe("HSCANARY");
    expect(got).not.toContain("HSCANARY");
  });
});

describe("text that only looks like a DOCTYPE", () => {
  it("is kept in a shared string written as CDATA", () => {
    const parts = xlsxParts();
    parts["xl/sharedStrings.xml"] = enc.encode(text(parts, "xl/sharedStrings.xml").replace(MARKER, "<![CDATA[<!DOCTYPE html>]]>"));

    expect((Agent.read(zipOf(parts)) as any).sheets[0].cells.A3.value).toBe("<!DOCTYPE html>");
  });

  it("does not stop a part being read from a comment", () => {
    const parts = xlsxParts();
    parts["xl/worksheets/sheet1.xml"] = enc.encode(withDoctype(text(parts, "xl/worksheets/sheet1.xml"), "<!-- <!DOCTYPE x> -->"));

    expect((Agent.read(zipOf(parts)) as any).sheets[0].cells.A1.value).toBe(MARKER);
  });
});

describe("the gate", () => {
  it("returns a tree for a clean part", () => {
    expect(parseXml('<?xml version="1.0"?><r a="1"><b>t</b></r>')?.attrs.a).toBe("1");
  });

  it("refuses a declaration of any kind, wherever it sits", () => {
    expect(parseXml(`${DOCTYPE}<r/>`)).toBeNull();
    expect(parseXml(`<r>${DOCTYPE}</r>`)).toBeNull();
    expect(parseXml("<!ELEMENT r ANY><r/>")).toBeNull();
  });

  it("finds a DOCTYPE past a comment that reads as ended once the encoding is ignored", () => {
    // U+2D2D is the bytes 2D 2D in UTF-16LE, so a part decoded as UTF-8 reads
    // "-->" here and the DOCTYPE after it sits outside the comment it is really in.
    const raw = concat([0xff, 0xfe], utf16le(`<!-- ⴭ> x -->${DOCTYPE}<r>&x;</r>`));
    const parts = xlsxParts();
    parts["xl/workbook.xml"] = raw;

    expect(() => Agent.read(zipOf(parts))).toThrow("failed to parse xl/workbook.xml");
  });

  it("leaves nothing in src/ decoding a part or scanning for a DOCTYPE outside the gate", () => {
    const root = join(import.meta.dirname, "..", "src");
    const offenders: string[] = [];
    const walk = (dir: string): void => {
      for (const name of readdirSync(dir)) {
        const path = join(dir, name);
        if (statSync(path).isDirectory()) {
          walk(path);
        } else if (name.endsWith(".ts") && path !== join(root, "reader", "xml.ts")) {
          const source = readFileSync(path, "utf8");
          if (/<!DOCTYPE/i.test(source) || /decoder\.decode\(files\[name\]/.test(source) || /UTF8\.decode/.test(source) || /new DOMParser/.test(source)) {
            offenders.push(name);
          }
        }
      }
    };
    walk(root);

    expect(offenders).toEqual([]);
  });
});
