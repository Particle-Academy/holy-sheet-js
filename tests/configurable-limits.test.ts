import { describe, it, expect, vi } from "vitest";
import { existsSync, readFileSync, writeFileSync, unlinkSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { deflateRawSync } from "node:zlib";
import { randomBytes } from "node:crypto";
import { Agent, ArchiveRefused, DEFAULT_ARCHIVE_LIMITS, unzipSync } from "../src";
import { crc32 } from "../src/zip/crc32";

/**
 * Every limit a read enforces can be raised by the app (owner, 2026-11-07).
 * For each archive limit: a file the DEFAULT refuses, refused by name, then
 * read once that one limit is raised. Mirrors PHP ConfigurableLimitsTest; the
 * small cases (lowering each limit) are the shared cases.json.
 */
const DIR = process.env.HOLY_SHEET_PHP_SRC
  ? join(process.env.HOLY_SHEET_PHP_SRC, "..", "tests", "fixtures", "read-options")
  : join(__dirname, "..", "..", "holy-sheet", "tests", "fixtures", "read-options");
if (!existsSync(join(DIR, "two-sheets.xlsx"))) throw new Error(`read-option fixtures not found at ${DIR}`);
const TWO_SHEETS = new Uint8Array(readFileSync(join(DIR, "two-sheets.xlsx")));
const MiB = 1048576;

type Member = { name: string; data: Uint8Array; deflate: boolean };

/** A zip of `members`: STORE or raw DEFLATE (level 9), no data descriptors. */
function zip(members: Member[]): Uint8Array {
  const enc = new TextEncoder();
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  for (const m of members) {
    const name = enc.encode(m.name);
    const body = m.deflate ? new Uint8Array(deflateRawSync(m.data, { level: 9 })) : m.data;
    const crc = crc32(m.data);
    const local = new Uint8Array(30 + name.length);
    const l = new DataView(local.buffer);
    l.setUint32(0, 0x04034b50, true); l.setUint16(4, 20, true); l.setUint16(6, 0x0800, true);
    l.setUint16(8, m.deflate ? 8 : 0, true); l.setUint16(12, 0x0021, true);
    l.setUint32(14, crc, true); l.setUint32(18, body.length, true); l.setUint32(22, m.data.length, true);
    l.setUint16(26, name.length, true); local.set(name, 30);
    const central = new Uint8Array(46 + name.length);
    const c = new DataView(central.buffer);
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
    c.setUint16(10, m.deflate ? 8 : 0, true); c.setUint16(14, 0x0021, true);
    c.setUint32(16, crc, true); c.setUint32(20, body.length, true); c.setUint32(24, m.data.length, true);
    c.setUint16(28, name.length, true); c.setUint32(42, offset, true); central.set(name, 46);
    locals.push(local, body); centrals.push(central);
    offset += local.length + body.length;
  }
  const centralSize = centrals.reduce((n, p) => n + p.length, 0);
  const end = new Uint8Array(22);
  const e = new DataView(end.buffer);
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, members.length, true); e.setUint16(10, members.length, true);
  e.setUint32(12, centralSize, true); e.setUint32(16, offset, true);
  const out = new Uint8Array(offset + centralSize + 22);
  let at = 0;
  for (const p of [...locals, ...centrals, end]) { out.set(p, at); at += p.length; }
  return out;
}

/** two-sheets.xlsx plus `extra`, as bytes. */
function archive(extra: Member[]): Uint8Array {
  const parts = unzipSync(TWO_SHEETS);
  return zip([...Object.entries(parts).map(([name, data]) => ({ name, data, deflate: true })), ...extra]);
}

/** `bytes` long, a 4 KiB pseudo-random block repeated: deflates about 100:1. */
function filler(bytes: number): Uint8Array {
  const block = new Uint8Array(4096);
  let s = 7;
  for (let i = 0; i < block.length; i++) { s = (s * 1103515245 + 12345) >>> 0; block[i] = s >>> 24; }
  const out = new Uint8Array(bytes);
  for (let at = 0; at < bytes; at += 4096) out.set(block.subarray(0, Math.min(4096, bytes - at)), at);
  return out;
}

function refusal(run: () => unknown): { reason: string; limit: string | null } | null {
  try { run(); } catch (e) {
    if (e instanceof ArchiveRefused) return { reason: e.reason, limit: e.limit };
    throw e;
  }
  return null;
}

async function asyncRefusal(run: () => Promise<unknown>): Promise<{ reason: string; limit: string | null } | null> {
  try { await run(); } catch (e) {
    if (e instanceof ArchiveRefused) return { reason: e.reason, limit: e.limit };
    throw e;
  }
  return null;
}

const sheets = (schema: any): string[] => schema.sheets.map((s: any) => s.name); // eslint-disable-line @typescript-eslint/no-explicit-any

describe("every limit is configurable", () => {
  it("raises maxEntries past a file the default refuses", () => {
    const extra: Member[] = [];
    for (let i = 0; i <= DEFAULT_ARCHIVE_LIMITS.maxEntries; i++) extra.push({ name: `pad/${i}`, data: new Uint8Array(), deflate: false });
    const bytes = archive(extra);
    expect(refusal(() => Agent.read(bytes))).toEqual({ reason: "archive_limit", limit: "maxEntries" });
    expect(sheets(Agent.read(bytes, { limits: { maxEntries: 20000 } }))).toEqual(["Data", "Summary"]);
  });

  it("raises maxCompressionRatio past a file the default refuses", () => {
    const bytes = archive([{ name: "pad.bin", data: new Uint8Array(4 * MiB), deflate: true }]);
    expect(refusal(() => Agent.read(bytes))).toEqual({ reason: "archive_limit", limit: "maxCompressionRatio" });
    expect(sheets(Agent.read(bytes, { limits: { maxCompressionRatio: 2000 } }))).toEqual(["Data", "Summary"]);
  });

  it("raises maxPartBytes past a file the default refuses", () => {
    const bytes = archive([{ name: "pad.bin", data: filler(65 * MiB), deflate: true }]);
    expect(refusal(() => Agent.read(bytes))).toEqual({ reason: "archive_limit", limit: "maxPartBytes" });
    expect(sheets(Agent.read(bytes, { limits: { maxPartBytes: 128 * MiB } }))).toEqual(["Data", "Summary"]);
  });

  it("raises maxExpandedBytes past a file the default refuses", () => {
    const data = filler(55 * MiB);
    const bytes = archive([1, 2, 3, 4, 5].map((i) => ({ name: `pad${i}.bin`, data, deflate: true })));
    expect(refusal(() => Agent.read(bytes))).toEqual({ reason: "archive_limit", limit: "maxExpandedBytes" });
    expect(sheets(Agent.read(bytes, { limits: { maxExpandedBytes: 512 * MiB } }))).toEqual(["Data", "Summary"]);
  });

  it("raises maxArchiveBytes past a file the default refuses, from bytes and from a path", async () => {
    const bytes = archive([
      { name: "a.bin", data: new Uint8Array(randomBytes(33 * MiB)), deflate: false },
      { name: "b.bin", data: new Uint8Array(randomBytes(33 * MiB)), deflate: false },
    ]);
    const path = join(mkdtempSync(join(tmpdir(), "hs-limits-")), "big.xlsx");
    writeFileSync(path, bytes);
    const raised = { limits: { maxArchiveBytes: 128 * MiB } };
    try {
      expect(refusal(() => Agent.read(bytes))).toEqual({ reason: "archive_limit", limit: "maxArchiveBytes" });
      expect(await asyncRefusal(() => Agent.describe(path))).toEqual({ reason: "archive_limit", limit: "maxArchiveBytes" });
      expect(sheets(Agent.read(bytes, raised))).toEqual(["Data", "Summary"]);
      expect(sheets(await Agent.describe(path, raised))).toEqual(["Data", "Summary"]);
    } finally {
      unlinkSync(path);
    }
  });

  it("applies every limit to describe() from a path, as read() does", async () => {
    const path = join(DIR, "two-sheets.xlsx");
    const lowered = { maxArchiveBytes: 1000, maxEntries: 5, maxPartBytes: 600, maxExpandedBytes: 1000, maxCompressionRatio: 2 };
    for (const [limit, value] of Object.entries(lowered)) {
      expect(await asyncRefusal(() => Agent.describe(path, { limits: { [limit]: value } }))).toEqual({ reason: "archive_limit", limit });
    }
  });

  it("refuses a part too long for one JavaScript string as a typed archive_limit, not a bare Error", () => {
    // V8 cannot hold a string over 0x1fffffe8 (536,870,888) characters, so a
    // worksheet part past that cannot be decoded at all. Measured: the 1,000,000
    // x 10 workbook (a 553,315,173-byte sheet part) threw a plain Error. The
    // decoder is made to throw V8's own error here, because a real part that
    // size costs a gigabyte to build in a unit test.
    const original = TextDecoder.prototype.decode;
    const spy = vi.spyOn(TextDecoder.prototype, "decode").mockImplementation(function (this: TextDecoder, ...args) {
      const text = original.apply(this, args);
      // Only the worksheet part is "too long"; member names and other parts decode as usual.
      if (text.includes("<sheetData")) throw new Error("Cannot create a string longer than 0x1fffffe8 characters");
      return text;
    });
    try {
      let caught: unknown = null;
      try { Agent.read(TWO_SHEETS); } catch (e) { caught = e; }
      expect(caught).toBeInstanceOf(ArchiveRefused);
      const e = caught as ArchiveRefused;
      expect([e.reason, e.limit, e.max]).toEqual(["archive_limit", "maxPartBytes", 536870888]);
      expect(e.message).toContain("longer than this JavaScript runtime can hold in one string");
    } finally {
      spy.mockRestore();
    }
  });

  it("refuses a bad limit in describe() before the file is read", async () => {
    await expect(Agent.describe(join(DIR, "two-sheets.xlsx"), { limits: { maxPartBytes: 0 } })).rejects.toThrow(
      "archive_limits (maxPartBytes must be a positive integer)",
    );
  });
});
