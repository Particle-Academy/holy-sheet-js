import { describe, it, expect } from "vitest";
import { deflateRawSync, gzipSync } from "node:zlib";
import { zipSync, unzipSync, inflateRaw, crc32 } from "../src/zip";

const enc = new TextEncoder();
const dec = new TextDecoder();

describe("zip writer/reader round-trip", () => {
  it("stores and recovers multiple files byte-for-byte", () => {
    const files = [
      { name: "[Content_Types].xml", data: enc.encode("<a/>") },
      { name: "xl/worksheets/sheet1.xml", data: enc.encode("<sheet>123 αβγ</sheet>") },
      { name: "deep/nested/big.bin", data: new Uint8Array(5000).map((_, i) => i % 256) },
    ];
    const zip = zipSync(files);
    // valid local-file-header signature at start
    expect(zip[0]).toBe(0x50);
    expect(zip[1]).toBe(0x4b);
    const out = unzipSync(zip);
    for (const f of files) {
      expect(dec.decode(out[f.name]!)).toBe(dec.decode(f.data));
      expect(Array.from(out[f.name]!)).toEqual(Array.from(f.data));
    }
  });
});

describe("inflateRaw vs node zlib", () => {
  it("inflates dynamic-Huffman deflate streams", () => {
    const text = "Hello, OOXML! ".repeat(500) + "the quick brown fox jumps over the lazy dog";
    const raw = deflateRawSync(Buffer.from(text));
    const back = inflateRaw(new Uint8Array(raw));
    expect(dec.decode(back)).toBe(text);
  });

  it("inflates highly-repetitive data (back-references)", () => {
    const data = new Uint8Array(20000).map((_, i) => (i * 7) % 13);
    const raw = deflateRawSync(Buffer.from(data));
    const back = inflateRaw(new Uint8Array(raw));
    expect(Array.from(back)).toEqual(Array.from(data));
  });

  it("reads a DEFLATE-compressed zip entry (external-style)", () => {
    // Hand-build a zip entry that uses method 8 to exercise the reader's inflate path.
    const payload = enc.encode("<x>" + "z".repeat(2000) + "</x>");
    const compressed = new Uint8Array(deflateRawSync(Buffer.from(payload)));
    const name = "ppt/slide.xml";
    const nb = enc.encode(name);
    const crc = crc32(payload);
    const lh = new Uint8Array(30 + nb.length + compressed.length);
    const dvw = new DataView(lh.buffer);
    dvw.setUint32(0, 0x04034b50, true);
    dvw.setUint16(4, 20, true);
    dvw.setUint16(6, 0x0800, true);
    dvw.setUint16(8, 8, true); // DEFLATE
    dvw.setUint32(14, crc, true);
    dvw.setUint32(18, compressed.length, true);
    dvw.setUint32(22, payload.length, true);
    dvw.setUint16(26, nb.length, true);
    lh.set(nb, 30);
    lh.set(compressed, 30 + nb.length);
    const cd = new Uint8Array(46 + nb.length);
    const cdv = new DataView(cd.buffer);
    cdv.setUint32(0, 0x02014b50, true);
    cdv.setUint16(6, 20, true);
    cdv.setUint16(8, 0x0800, true);
    cdv.setUint16(10, 8, true);
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, compressed.length, true);
    cdv.setUint32(24, payload.length, true);
    cdv.setUint16(28, nb.length, true);
    cdv.setUint32(42, 0, true);
    cd.set(nb, 46);
    const eocd = new Uint8Array(22);
    const edv = new DataView(eocd.buffer);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(8, 1, true);
    edv.setUint16(10, 1, true);
    edv.setUint32(12, cd.length, true);
    edv.setUint32(16, lh.length, true);
    const zip = new Uint8Array(lh.length + cd.length + eocd.length);
    zip.set(lh, 0);
    zip.set(cd, lh.length);
    zip.set(eocd, lh.length + cd.length);

    const out = unzipSync(zip);
    expect(dec.decode(out[name]!)).toBe(dec.decode(payload));
  });

  // keep gzipSync import used (sanity that node zlib is present)
  it("node zlib present", () => {
    expect(gzipSync(Buffer.from("x")).length).toBeGreaterThan(0);
  });
});
