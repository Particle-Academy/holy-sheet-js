import { describe, expect, it } from "vitest";
import { deflateRawSync, inflateRawSync } from "node:zlib";
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, extname, join } from "node:path";
import { randomBytes } from "node:crypto";
import { Agent, ArchiveRefused, OfficeTemplateRefused, UnsupportedFormatException, XlsxReader } from "../src/index";
import { crc32 } from "../src/zip/crc32";
import { run } from "../src/cli-run";

/*
 * Official template formats: Excel .xltx and LibreOffice .ots. Mirrors PHP
 * tests/Unit/OfficeTemplateTest.php.
 *
 * A template IS its document twin with a different declared type: the main
 * part's content type in [Content_Types].xml (OOXML), or the `mimetype` entry
 * plus the manifest's root entry (ODF 1.3 Part 2 section 3.3 requires the two to
 * be equal). read() marks `sourceKind: "template"`; the converter flips the
 * declaration and copies every other part byte for byte.
 *
 * Fixtures are zipped and unzipped here with node:zlib, independently of the
 * production archive code, so a converter bug cannot also forge its evidence.
 */

const DOC = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const TEMPLATE = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
const MACRO_TEMPLATE = "application/vnd.ms-excel.template.macroEnabled.main+xml";
const MACRO_DOC = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
const PPT_TEMPLATE = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
const ODS = "application/vnd.oasis.opendocument.spreadsheet";
const OTS = "application/vnd.oasis.opendocument.spreadsheet-template";

const enc = new TextEncoder();
const dec = new TextDecoder();
type Entries = Map<string, Uint8Array>;
type Raw = { name: string; method: number; data: Uint8Array; localExtra: number };

const schema = {
  sheets: [{ name: "Budget", columns: [{ header: "Item" }, { header: "Cost" }], rows: [["Rent", 1200], ["Power", 85]] }],
};

/** Independent central-directory walk: entry name, method, uncompressed data. */
function rawEntries(bytes: Uint8Array): Raw[] {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.length - 22;
  while (end >= 0 && v.getUint32(end, true) !== 0x06054b50) end--;
  const count = v.getUint16(end + 10, true);
  let at = v.getUint32(end + 16, true);
  const out: Raw[] = [];
  for (let i = 0; i < count; i++) {
    const method = v.getUint16(at + 10, true), compressed = v.getUint32(at + 20, true);
    const nameLength = v.getUint16(at + 28, true), extra = v.getUint16(at + 30, true), comment = v.getUint16(at + 32, true);
    const local = v.getUint32(at + 42, true);
    const name = dec.decode(bytes.subarray(at + 46, at + 46 + nameLength));
    const localExtra = v.getUint16(local + 28, true);
    const start = local + 30 + v.getUint16(local + 26, true) + localExtra;
    const body = bytes.subarray(start, start + compressed);
    out.push({ name, method, localExtra, data: method === 8 ? new Uint8Array(inflateRawSync(body)) : body.slice() });
    at += 46 + nameLength + extra + comment;
  }
  return out;
}

function entries(bytes: Uint8Array): Entries {
  return new Map(rawEntries(bytes).map((e) => [e.name, e.data]));
}

/** Zip name => bytes in order; `mimetype` STORED (as ODF requires), the rest deflated. */
function zip(files: Entries): Uint8Array {
  const locals: Uint8Array[] = [], central: Uint8Array[] = [];
  let offset = 0;
  for (const [entryName, raw] of files) {
    const name = enc.encode(entryName), method = entryName === "mimetype" ? 0 : 8;
    const data = method === 8 ? new Uint8Array(deflateRawSync(raw)) : raw;
    const local = new Uint8Array(30 + name.length + data.length), lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true); lv.setUint16(8, method, true);
    lv.setUint16(12, 33, true); lv.setUint32(14, crc32(raw) >>> 0, true); lv.setUint32(18, data.length, true);
    lv.setUint32(22, raw.length, true); lv.setUint16(26, name.length, true); local.set(name, 30); local.set(data, 30 + name.length);
    const cd = new Uint8Array(46 + name.length), cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, method, true); cv.setUint16(14, 33, true); cv.setUint32(16, crc32(raw) >>> 0, true);
    cv.setUint32(20, data.length, true); cv.setUint32(24, raw.length, true); cv.setUint16(28, name.length, true);
    cv.setUint32(42, offset, true); cd.set(name, 46);
    locals.push(local); central.push(cd); offset += local.length;
  }
  const centralLength = central.reduce((n, c) => n + c.length, 0);
  const eocd = new Uint8Array(22), ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.size, true); ev.setUint16(10, files.size, true);
  ev.setUint32(12, centralLength, true); ev.setUint32(16, offset, true);
  const out = new Uint8Array(offset + centralLength + 22);
  let pos = 0;
  for (const part of [...locals, ...central, eocd]) { out.set(part, pos); pos += part.length; }
  return out;
}

const xlsx = (): Uint8Array => Agent.toBytes(schema);

/** The xlsx with its main part re-declared; nothing else differs. */
function retyped(contentType: string, extra: Record<string, Uint8Array> = {}, base: Uint8Array = xlsx()): Uint8Array {
  const files = entries(base);
  files.set("[Content_Types].xml", enc.encode(dec.decode(files.get("[Content_Types].xml")!).replace(DOC, contentType)));
  for (const [name, data] of Object.entries(extra)) files.set(name, data);
  return zip(files);
}

const xltx = (): Uint8Array => retyped(TEMPLATE);

function odf(mimetype: string, body = "spreadsheet", manifestType: string = mimetype): Uint8Array {
  const inner = body === "spreadsheet"
    ? '<table:table table:name="Budget"><table:table-row>'
      + '<table:table-cell office:value-type="string"><text:p>Rent</text:p></table:table-cell>'
      + '<table:table-cell office:value-type="float" office:value="1200"><text:p>1200</text:p></table:table-cell>'
      + "</table:table-row></table:table>"
    : "<text:p>Not a spreadsheet</text:p>";
  return zip(new Map([
    ["mimetype", enc.encode(mimetype)],
    ["content.xml", enc.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
      + ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3">'
      + `<office:body><office:${body}>${inner}</office:${body}></office:body></office:document-content>`)],
    ["Thumbnails/opaque.bin", new Uint8Array([0, 1, 2, 98, 255])],
    ["META-INF/manifest.xml", enc.encode('<?xml version="1.0" encoding="UTF-8"?>\n'
      + '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">\n'
      + ` <manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${manifestType}"/>\n`
      + ' <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>\n'
      + "</manifest:manifest>")],
  ]));
}

function refused(action: () => unknown, reason: string): void {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(OfficeTemplateRefused);
    expect((error as OfficeTemplateRefused).reason).toBe(reason);
    return;
  }
  throw new Error(`Expected a ${reason} refusal; the input was accepted.`);
}

function archiveRefused(action: () => unknown, reason: string): void {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(ArchiveRefused);
    expect((error as ArchiveRefused).reason).toBe(reason);
    return;
  }
  throw new Error(`Expected archive refusal ${reason}; the input was accepted.`);
}

function without(read: Record<string, unknown>): Record<string, unknown> {
  const { sourceKind: _, ...rest } = read;
  return rest;
}

describe("reading templates", () => {
  it("reads an Excel .xltx and marks it as a template", () => {
    const base = xlsx();
    const template = Agent.read(retyped(TEMPLATE, {}, base));
    expect(template["sourceKind"]).toBe("template");
    expect(without(template)).toEqual(Agent.read(base));
  });

  it("leaves an .xlsx read exactly as it was: no sourceKind key", () => {
    expect(Agent.read(xlsx())).not.toHaveProperty("sourceKind");
  });

  it("reads a LibreOffice .ots and marks it as a template", () => {
    const document = Agent.read(odf(ODS));
    const template = Agent.read(odf(OTS));
    expect(document).not.toHaveProperty("sourceKind");
    expect(template["sourceKind"]).toBe("template");
    expect(without(template)).toEqual(document);
  });

  it("marks the Workbook itself, for readWorkbook() callers", () => {
    expect(new XlsxReader().readWorkbook(xltx()).sourceKind).toBe("template");
    expect(new XlsxReader().readWorkbook(xlsx()).sourceKind).toBe("document");
  });

  it("keeps a template read valid and writable", () => {
    const template = Agent.read(xltx());
    expect(Agent.validate(template)).toEqual([]);
    expect(Agent.read(Agent.toBytes(template))).not.toHaveProperty("sourceKind");
  });

  it("refuses a macro-enabled .xltm by name rather than reading around the VBA", () => {
    refused(() => Agent.read(retyped(MACRO_TEMPLATE)), "macro_enabled");
  });

  it("refuses an .xltx that carries a VBA project anyway", () => {
    refused(() => Agent.read(retyped(TEMPLATE, { "xl/vbaProject.bin": enc.encode("not really") })), "macro_enabled");
  });

  it("still reads a macro-enabled WORKBOOK (.xlsm) as before: only templates changed", () => {
    expect(Agent.read(retyped(MACRO_DOC))).not.toHaveProperty("sourceKind");
  });
});

describe("hostile templates", () => {
  it("refuses a template whose content type claims another family", () => {
    refused(() => Agent.read(retyped(PPT_TEMPLATE)), "family_mismatch");

    let files = entries(xlsx());
    files.set("[Content_Types].xml", enc.encode(dec.decode(files.get("[Content_Types].xml")!)
      .replace("</Types>", `<Override PartName="/ppt/presentation.xml" ContentType="${PPT_TEMPLATE}"/></Types>`)));
    const beside = zip(files);
    refused(() => Agent.read(beside), "family_mismatch");

    files = entries(xlsx());
    files.set("[Content_Types].xml", enc.encode(dec.decode(files.get("[Content_Types].xml")!)
      .replace('ContentType="application/xml"', `ContentType="${TEMPLATE}"`)));
    const viaDefault = zip(files);
    refused(() => Agent.read(viaDefault), "family_mismatch");
  });

  it("refuses an ODF spreadsheet template whose body is text", () => {
    refused(() => Agent.read(odf(OTS, "text")), "family_mismatch");
  });

  it("refuses an ODF template of another family by name", () => {
    expect(() => Agent.read(odf("application/vnd.oasis.opendocument.text-template", "text"))).toThrow(UnsupportedFormatException);
  });

  it("refuses an ODF template whose manifest names another family", () => {
    refused(() => Agent.read(odf(OTS, "spreadsheet", "application/vnd.oasis.opendocument.text-template")), "family_mismatch");
  });

  it("refuses a template zip bomb through the archive guard, reading and converting", () => {
    const bomb = retyped(TEMPLATE, { "xl/media/bomb.bin": new Uint8Array(65 * 1024 * 1024) });
    archiveRefused(() => Agent.read(bomb), "archive_limit");
    archiveRefused(() => Agent.convertTemplate(bomb), "archive_limit");
  });

  it("refuses a template with a traversing entry name, reading and converting", () => {
    const evil = retyped(TEMPLATE, { "../evil.xml": enc.encode("<x/>") });
    archiveRefused(() => Agent.read(evil), "archive_path");
    archiveRefused(() => Agent.convertTemplate(evil), "archive_path");
  });
});

describe("converting templates", () => {
  it("converts an .xltx to an .xlsx changing only the content-type value", () => {
    const template = xltx();
    const document = Agent.convertTemplate(template, "document");
    const before = entries(template), after = entries(document);
    expect([...after.keys()]).toEqual([...before.keys()]);
    for (const [name, data] of before) {
      if (name === "[Content_Types].xml") {
        expect(dec.decode(after.get(name)!)).toBe(dec.decode(data).replace(TEMPLATE, DOC));
      } else {
        expect(after.get(name), `${name} changed`).toEqual(data);
      }
    }
    expect(Agent.read(document)).not.toHaveProperty("sourceKind");
  });

  it("round-trips xlsx to xltx and back with every part byte-identical", () => {
    const document = xlsx();
    const template = Agent.convertTemplate(document, "template");
    const back = Agent.convertTemplate(template, "document");
    expect(Agent.read(template)["sourceKind"]).toBe("template");
    expect(entries(back)).toEqual(entries(document));

    // Untouched parts are copied as RECORDS, not re-encoded: same compression.
    const methods = (bytes: Uint8Array) => rawEntries(bytes).filter((e) => e.name !== "[Content_Types].xml").map((e) => [e.name, e.method]);
    expect(methods(template)).toEqual(methods(document));
  });

  it("converts a Node Buffer that is a view into a larger allocation", () => {
    // readFileSync hands back Buffers, and Buffer#slice is a VIEW: an offset
    // write through `record.buffer` once landed in the wrong bytes entirely.
    const bytes = xltx();
    const view = Buffer.concat([Buffer.alloc(7), Buffer.from(bytes)]).subarray(7);
    expect(view.byteOffset).toBeGreaterThan(0);
    expect(Agent.read(Agent.convertTemplate(view, "document"))).not.toHaveProperty("sourceKind");
    expect(Buffer.from(view).equals(Buffer.from(bytes))).toBe(true); // the input is untouched
  });

  it("flips direction when no target is given", () => {
    expect(Agent.read(Agent.convertTemplate(xlsx()))["sourceKind"]).toBe("template");
    expect(Agent.read(Agent.convertTemplate(xltx()))).not.toHaveProperty("sourceKind");
  });

  it("converts .ods to .ots keeping mimetype first, STORED and without an extra field", () => {
    const ods = odf(ODS);
    const ots = Agent.convertTemplate(ods, "template");
    const before = entries(ods), after = entries(ots);
    const first = rawEntries(ots)[0]!;
    expect([...after.keys()]).toEqual([...before.keys()]);
    expect(first.name).toBe("mimetype");
    expect(first.method).toBe(0);
    expect(first.localExtra).toBe(0);
    expect(dec.decode(ots.subarray(30, 38))).toBe("mimetype");
    expect(dec.decode(ots.subarray(38, 38 + OTS.length))).toBe(OTS);

    // ODF 1.3 Part 2 section 3.3: mimetype SHALL equal the manifest's root
    // media-type, so that one attribute value moves with it and nothing else.
    expect(dec.decode(after.get("META-INF/manifest.xml")!)).toBe(dec.decode(before.get("META-INF/manifest.xml")!).replace(
      `manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODS}"`,
      `manifest:full-path="/" manifest:version="1.3" manifest:media-type="${OTS}"`,
    ));
    for (const [name, data] of before) {
      if (name !== "mimetype" && name !== "META-INF/manifest.xml") expect(after.get(name), `${name} changed`).toEqual(data);
    }
    expect(Agent.read(ots)["sourceKind"]).toBe("template");
    expect(entries(Agent.convertTemplate(ots, "document"))).toEqual(before);
  });

  it("refuses conversions it must not make", () => {
    refused(() => Agent.convertTemplate(xltx(), "template"), "already_template");
    refused(() => Agent.convertTemplate(xlsx(), "document"), "already_document");
    refused(() => Agent.convertTemplate(retyped(MACRO_TEMPLATE), "document"), "macro_enabled");
    refused(() => Agent.convertTemplate(retyped(MACRO_DOC), "template"), "macro_enabled");
    refused(() => Agent.convertTemplate(xlsx(), "presentation" as never), "bad_target");
    refused(() => Agent.convertTemplate(retyped(PPT_TEMPLATE)), "family_mismatch");
  });

  it("refuses an ODF package whose mimetype is not stored first", () => {
    const files = entries(odf(ODS));
    const mimetype = files.get("mimetype")!;
    files.delete("mimetype");
    files.set("mimetype", mimetype); // re-added LAST (still stored)
    const late = zip(files);
    refused(() => Agent.convertTemplate(late), "odf_mimetype");
  });

  it("refuses a content-type declaration it cannot rewrite unambiguously", () => {
    const files = entries(xltx());
    files.set("[Content_Types].xml", enc.encode(dec.decode(files.get("[Content_Types].xml")!).replace(
      '<Override PartName="/xl/workbook.xml"',
      `<!-- <Override PartName="/xl/workbook.xml" ContentType="${TEMPLATE}"/> --><Override PartName="/xl/workbook.xml"`,
    )));
    const decoy = zip(files);
    refused(() => Agent.convertTemplate(decoy), "not_convertible");
  });

  it("names a word-processing package instead of converting it", () => {
    const docx = zip(new Map([["[Content_Types].xml", enc.encode("<Types/>")], ["word/document.xml", enc.encode("<w:document/>")]]));
    expect(() => Agent.convertTemplate(docx)).toThrow(UnsupportedFormatException);
    expect(() => Agent.convertTemplate(odf("application/vnd.oasis.opendocument.text", "text"))).toThrow(UnsupportedFormatException);
  });
});

describe("template:convert CLI", () => {
  it("converts, refuses and explains", async () => {
    const dir = mkdtempSync(join(tmpdir(), "hst-cli-"));
    const out: string[] = [], err: string[] = [];
    const io = { stdout: (t: string) => { out.push(t); }, stderr: (t: string) => { err.push(t); } };
    try {
      writeFileSync(join(dir, "in.xltx"), xltx());
      expect(await run(["template:convert", join(dir, "in.xltx"), join(dir, "out.xlsx")], io)).toBe(0);
      expect(out.join("")).toContain("document");
      expect(Agent.read(readFileSync(join(dir, "out.xlsx")))).not.toHaveProperty("sourceKind");

      expect(await run(["template:convert", join(dir, "out.xlsx"), join(dir, "out.xlsx"), "--to=template"], io)).toBe(1);
      expect(err.join("")).toContain("exists");

      expect(await run(["template:convert", join(dir, "out.xlsx"), join(dir, "back.xltx"), "--to=template"], io)).toBe(0);
      expect(Agent.read(readFileSync(join(dir, "back.xltx")))["sourceKind"]).toBe("template");

      writeFileSync(join(dir, "macro.xltm"), retyped(MACRO_TEMPLATE));
      expect(await run(["template:convert", join(dir, "macro.xltm"), join(dir, "macro.xlsx")], io)).toBe(1);
      expect(err.join("")).toContain("macro_enabled");
      expect(existsSync(join(dir, "macro.xlsx"))).toBe(false);

      expect(await run(["template:convert"], io)).toBe(2);
      expect(err.join("")).toContain("Usage");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ─── LibreOffice as an oracle (skipped where it is not installed) ─────────

const soffice = [process.env["LIBREOFFICE_BIN"], "C:/Program Files/LibreOffice/program/soffice.com", "/usr/bin/soffice",
  "/usr/local/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"].find((p) => p !== undefined && p !== "" && existsSync(p));

/** Headless convert; the profile lives in a SHORT temp path (long ones crash LO on Windows). */
function libreOffice(input: string, format: string, outDir: string): string {
  const profile = join(tmpdir(), `hjlo${randomBytes(3).toString("hex")}`).split("\\").join("/");
  const result = spawnSync(soffice!, [`-env:UserInstallation=file:///${profile.replace(/^\/+/, "")}`, "--headless", "--norestore",
    "--convert-to", format, "--outdir", outDir, input], { encoding: "utf8", timeout: 120_000, windowsHide: true });
  const out = join(outDir, `${basename(input, extname(input))}.${format.split(":")[0]}`);
  expect(result.status).toBe(0);
  expect(existsSync(out), `LibreOffice produced no ${format}`).toBe(true);
  return out;
}

describe.skipIf(soffice === undefined)("LibreOffice oracle", () => {
  it("reads what LibreOffice writes as .xltx and .ots, and LibreOffice reads our conversions", () => {
    const dir = mkdtempSync(join(tmpdir(), "hjo"));
    try {
      writeFileSync(join(dir, "src.xlsx"), xlsx());
      for (const [template, document] of [["xltx", "xlsx"], ["ots", "ods"]] as const) {
        const loTemplate = libreOffice(join(dir, "src.xlsx"), template, dir);
        const read = Agent.read(readFileSync(loTemplate));
        expect(read["sourceKind"], `LibreOffice .${template} not marked`).toBe("template");
        expect(JSON.stringify(read)).toContain("Rent");

        mkdirSync(join(dir, template));
        const loBack = libreOffice(loTemplate, document, join(dir, template));
        expect(Agent.read(readFileSync(loBack))).not.toHaveProperty("sourceKind");

        writeFileSync(join(dir, `ours.${document}`), Agent.convertTemplate(readFileSync(loTemplate), "document"));
        mkdirSync(join(dir, `ours-${template}`));
        const csv = libreOffice(join(dir, `ours.${document}`), "csv", join(dir, `ours-${template}`));
        expect(readFileSync(csv, "utf8")).toContain("Rent,1200");
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 120_000);
});
