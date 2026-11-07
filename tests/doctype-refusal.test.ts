import { mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { Agent, OfficeTemplateRefused, UnsupportedFormatException, unzipSync, zipSync } from "../src/index";
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
const SHEET_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const MACRO_TEMPLATE = "application/vnd.ms-excel.template.macroEnabled.main+xml";

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
            A3: 2,
          },
          charts: [{ type: "column", title: MARKER, series: [{ values: "A2:A3" }] }],
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
const MARKER_PARTS = ["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "docProps/core.xml", "xl/comments1.xml", "xl/charts/chart1.xml"];

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
  // A1 value, comment, A3 shared string, chart title and creator. Styles has its
  // own test: this port returns no format code as text.
  expect(count(readText(xlsxParts()), MARKER)).toBe(5);
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
    ["xl/drawings/_rels/drawing1.xml.rels", "../charts/chart1.xml", '"charts"'],
  ])("does not follow %s", (part, target, probe) => {
    const parts = xlsxParts();
    expect(text(parts, part)).toContain(target);
    parts[part] = enc.encode(withDoctype(text(parts, part)));

    expect(readText(parts)).not.toContain(probe);
  });

  it("does not read a drawing with one", () => {
    const parts = xlsxParts();
    parts["xl/drawings/drawing1.xml"] = enc.encode(withDoctype(text(parts, "xl/drawings/drawing1.xml")));

    expect(readText(parts)).not.toContain('"charts"');
  });
});

describe("what a consumer sees, per part: the same table as PHP", () => {
  // A refused part used to read as an absent one, so a worksheet carrying a
  // DOCTYPE came back as an EMPTY sheet and the caller could not tell that from
  // a sheet with nothing on it. A refusal now refuses the file, as one in
  // [Content_Types].xml always has. Mirrors PHP 2.11.5.
  const REFUSING_PARTS = [
    ...MARKER_PARTS,
    "xl/styles.xml",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels",
  ];

  it.each(REFUSING_PARTS.flatMap((p) => ENCODING_NAMES.map((e) => [p, e] as const)))(
    "refuses the file when %s carries a DOCTYPE, %s",
    (part, encoding) => {
      const parts = xlsxParts();
      parts[part] = ENCODINGS[encoding]!(withDoctype(text(parts, part)));

      expect(() => Agent.read(zipOf(parts))).toThrow(UnsupportedFormatException);
      expect(() => Agent.read(zipOf(parts))).toThrow("carries a DOCTYPE");
    },
  );

  it("refuses a worksheet with a DOCTYPE instead of reading it as an empty sheet", () => {
    const parts = xlsxParts();
    parts["xl/worksheets/sheet1.xml"] = enc.encode(withDoctype(text(parts, "xl/worksheets/sheet1.xml")));

    expect(() => Agent.read(zipOf(parts))).toThrow(UnsupportedFormatException);
  });

  it("still reads a file whose optional part is missing", () => {
    const parts = xlsxParts();
    delete parts["docProps/core.xml"];
    const read = Agent.read(zipOf(parts)) as any;

    expect(read.sheets[0].cells.A1.value).toBe(MARKER);
    expect(read.meta).toBeUndefined();
  });

  it("still reads a malformed worksheet as an empty sheet, because only a refusal raises", () => {
    const parts = xlsxParts();
    parts["xl/worksheets/sheet1.xml"] = enc.encode("<worksheet><sheetData>");
    const sheet = (Agent.read(zipOf(parts)) as any).sheets[0];

    expect(sheet.name).toBe("Sheet1");
    expect(sheet.cells ?? {}).toEqual({});
  });

  it("still throws for a malformed workbook part", () => {
    const parts = xlsxParts();
    parts["xl/workbook.xml"] = enc.encode("<?xml version='1.0'?>");

    expect(() => Agent.read(zipOf(parts))).toThrow("failed to parse xl/workbook.xml");
  });
});

describe("the subset is never read as content", () => {
  // An entity value may hold a '>', which is where a scan for the end of the
  // declaration would stop. What follows it must not become the part's root.
  const smuggling = (root: string): string => `<!DOCTYPE r [<!ENTITY y "a>${root}">]>`;

  it("takes no root from inside the subset", () => {
    expect(() => parseXml(`<?xml version="1.0"?>${smuggling("<evil/>")}<r/>`)).toThrow(UnsupportedFormatException);
  });

  it("takes no cells from inside the subset", () => {
    const parts = xlsxParts();
    const sheet = "<worksheet><sheetData><row r='1'><c r='A1' t='str'><v>HSSMUGGLED</v></c></row></sheetData></worksheet>";
    parts["xl/worksheets/sheet1.xml"] = enc.encode(withDoctype(text(parts, "xl/worksheets/sheet1.xml"), smuggling(sheet)));

    expect(readText(parts)).not.toContain("HSSMUGGLED");
  });
});

describe("the package declaration", () => {
  const macroTemplate = (): Parts => {
    const parts = xlsxParts();
    expect(text(parts, "[Content_Types].xml")).toContain(SHEET_MAIN);
    parts["[Content_Types].xml"] = enc.encode(text(parts, "[Content_Types].xml").replace(SHEET_MAIN, MACRO_TEMPLATE));
    return parts;
  };

  it("refuses a macro template for its declaration, so the cases below have something to skip", () => {
    expect(() => Agent.read(zipOf(macroTemplate()))).toThrow(OfficeTemplateRefused);
  });

  it.each(ENCODING_NAMES)("reads a clean declaration, %s", (encoding) => {
    const parts = xlsxParts();
    const clean = readText(parts);
    parts["[Content_Types].xml"] = ENCODINGS[encoding]!(text(parts, "[Content_Types].xml"));

    expect(readText(parts)).toBe(clean);
  });

  it.each(ENCODING_NAMES)("refuses the macro template when its declaration carries a DOCTYPE, %s", (encoding) => {
    const parts = macroTemplate();
    parts["[Content_Types].xml"] = ENCODINGS[encoding]!(withDoctype(text(parts, "[Content_Types].xml")));

    expect(() => Agent.read(zipOf(parts))).toThrow(UnsupportedFormatException);
  });

  it.each(ENCODING_NAMES)("refuses a plain workbook whose declaration carries a DOCTYPE, %s", (encoding) => {
    const parts = xlsxParts();
    parts["[Content_Types].xml"] = ENCODINGS[encoding]!(withDoctype(text(parts, "[Content_Types].xml")));

    expect(() => Agent.read(zipOf(parts))).toThrow(UnsupportedFormatException);
  });

  it("refuses a declaration this port cannot decode, and so cannot vouch for", () => {
    const parts = xlsxParts();
    parts["[Content_Types].xml"] = enc.encode(
      '<?xml version="1.0" encoding="UTF-7"?>+ADw-+ACE-DOCTYPE r +AFs-+ADw-+ACE-ENTITY x +ACI-HSINT+ACI-+AD4-+AF0-+AD4-+ADw-Types/+AD4-',
    );

    expect(() => Agent.read(zipOf(parts))).toThrow(UnsupportedFormatException);
  });

  it("refuses the conversion too", () => {
    const parts = xlsxParts();
    parts["[Content_Types].xml"] = enc.encode(withDoctype(text(parts, "[Content_Types].xml")));

    expect(() => Agent.convertTemplate(zipOf(parts))).toThrow(UnsupportedFormatException);
  });

  it("does not refuse a declaration for text that only looks like a DOCTYPE", () => {
    const parts = xlsxParts();
    const clean = readText(parts);
    parts["[Content_Types].xml"] = enc.encode(
      withDoctype(text(parts, "[Content_Types].xml"), "<!-- <!DOCTYPE x> -->").replace("</Types>", "<!-- <!DOCTYPE x> --></Types>"),
    );

    expect(readText(parts)).toBe(clean);
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

  it("is kept in a chart title written as CDATA", () => {
    const parts = xlsxParts();
    parts["xl/charts/chart1.xml"] = enc.encode(text(parts, "xl/charts/chart1.xml").split(MARKER).join("<![CDATA[<!DOCTYPE html>]]>"));

    expect((Agent.read(zipOf(parts)) as any).sheets[0].charts[0].title).toBe("<!DOCTYPE html>");
  });

  it("does not stop a part being read from a comment", () => {
    const parts = xlsxParts();
    parts["xl/worksheets/sheet1.xml"] = enc.encode(withDoctype(text(parts, "xl/worksheets/sheet1.xml"), "<!-- <!DOCTYPE x> -->"));

    expect((Agent.read(zipOf(parts)) as any).sheets[0].cells.A1.value).toBe(MARKER);
  });
});

describe("ods", () => {
  const ns =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"';
  const decl = '<?xml version="1.0" encoding="UTF-8"?>';
  const odsParts = (body: string = MARKER): Record<string, string> => ({
    mimetype: "application/vnd.oasis.opendocument.spreadsheet",
    "content.xml": `${decl}<office:document-content ${ns}><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell office:value-type="string"><text:p>${body}</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    "styles.xml": `${decl}<office:document-styles ${ns}/>`,
    "meta.xml": `${decl}<office:document-meta ${ns}><office:meta><meta:initial-creator>${body}</meta:initial-creator></office:meta></office:document-meta>`,
  });
  const packed = (parts: Record<string, string | Uint8Array>): Uint8Array =>
    zipOf(Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, typeof v === "string" ? enc.encode(v) : v])));

  it("reads a clean file", () => {
    expect(count(JSON.stringify(Agent.read(packed(odsParts()))), MARKER)).toBe(2);
  });

  it.each(["content.xml", "styles.xml", "meta.xml"].flatMap((p) => ENCODING_NAMES.map((e) => [p, e] as const)))(
    "refuses the file for a %s with a DOCTYPE, %s",
    (part, encoding) => {
      const parts: Record<string, string | Uint8Array> = odsParts();
      parts[part] = ENCODINGS[encoding]!(withDoctype(parts[part] as string));

      expect(() => Agent.read(packed(parts))).toThrow(UnsupportedFormatException);
      expect(() => Agent.read(packed(parts))).toThrow("carries a DOCTYPE");
    },
  );

  it.each(ENCODING_NAMES)("takes nothing from a refused meta part, %s", (encoding) => {
    const parts: Record<string, string | Uint8Array> = odsParts();
    parts["meta.xml"] = ENCODINGS[encoding]!(withDoctype((parts["meta.xml"] as string).split(MARKER).join("&x;")));
    let got: string;
    try {
      got = JSON.stringify(Agent.read(packed(parts)));
    } catch {
      got = "(refused)";
    }

    expect(got).not.toContain("&x;");
    expect(got).not.toContain(ENTITY);
  });
});

describe("the gate", () => {
  it("returns a tree for a clean part", () => {
    expect(parseXml('<?xml version="1.0"?><r a="1"><b>t</b></r>')?.attrs.a).toBe("1");
  });

  it("refuses a declaration of any kind, wherever it sits", () => {
    expect(() => parseXml(`${DOCTYPE}<r/>`)).toThrow(UnsupportedFormatException);
    expect(() => parseXml(`<?xml version="1.0"?><!-- c -->${DOCTYPE}<r/>`)).toThrow(UnsupportedFormatException);
    expect(() => parseXml(`<!doctype r><r/>`)).toThrow(UnsupportedFormatException);
    // Not a DOCTYPE before the root: PHP's libxml and Python's expat call these
    // malformed, and so a part read as absent, not a refused file.
    expect(parseXml(`<r>${DOCTYPE}</r>`)).toBeNull();
    expect(parseXml("<!ELEMENT r ANY><r/>")).toBeNull();
  });

  it("finds a DOCTYPE past a comment that reads as ended once the encoding is ignored", () => {
    const raw = concat([0xff, 0xfe], utf16le(`<!-- ⴭ> x -->${DOCTYPE}<r>&x;</r>`));
    const parts = xlsxParts();
    parts["xl/workbook.xml"] = raw;

    expect(() => Agent.read(zipOf(parts))).toThrow(UnsupportedFormatException);
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
