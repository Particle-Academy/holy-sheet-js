import type { XmlNode } from "../xml";

/**
 * OpenDocument namespace URIs, and lookups by URI. Mirrors PHP `Reader\Ods\Ns`.
 *
 * Elements and attributes are matched by URI, never by prefix: `table:` is a
 * convention every producer happens to follow, not something the format
 * promises. Parse with `parseXml(src, { namespaces: true })`.
 */
export const NS = {
  OFFICE: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  TABLE: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  TEXT: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  STYLE: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  FO: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  NUMBER: "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
  DC: "http://purl.org/dc/elements/1.1/",
  META: "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
  /** LibreOffice's extension namespace; carries `value-type="error"` for a formula error. */
  CALCEXT: "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
} as const;

/** The first child element with this namespace and local name. */
export function child(node: XmlNode | null | undefined, ns: string, name: string): XmlNode | undefined {
  return node?.children.find((c) => c.ns === ns && c.name === name);
}

/** Child elements in this namespace, in document order. */
export function childrenIn(node: XmlNode | null | undefined, ns: string): XmlNode[] {
  return node ? node.children.filter((c) => c.ns === ns) : [];
}

/** A namespaced attribute, or undefined when absent. */
export function attr(node: XmlNode | null | undefined, ns: string, name: string): string | undefined {
  return node?.nsAttrs?.[`{${ns}}${name}`];
}

/** An element's own text: its direct text runs, as SimpleXML's `(string)` cast reads it. */
export function ownText(node: XmlNode | null | undefined): string {
  return node ? node.text : "";
}

// ─── PHP semantics the reference reader relies on ──────────────────────────
//
// The PHP reader is the reference, and several of its primitives are narrower
// than their JavaScript look-alikes. `trim` and `strtolower` are ASCII-only;
// `(int)` reads a leading integer and gives 0 for anything else. Using the
// JavaScript versions would agree on every fixture and disagree on the first
// file with a narrow no-break space in a currency format.

/** PHP `trim()`: space, tab, LF, CR, NUL and vertical tab only. */
export function phpTrim(s: string): string {
  return s.replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

/** PHP 8 `strtolower()`: ASCII letters only. */
export function asciiLower(s: string): string {
  return s.replace(/[A-Z]/g, (c) => c.toLowerCase());
}

/** PHP `(int) $string`: the leading integer, 0 when there is none. */
export function phpInt(s: string): number {
  const m = /^[ \t\n\r\v\f]*([+-]?\d+)/.exec(s);
  return m ? parseInt(m[1]!, 10) : 0;
}

/** PHP 8 `is_numeric()` on a string. */
export function phpIsNumeric(s: string): boolean {
  return /^[ \t\n\r\v\f]*[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?[ \t\n\r\v\f]*$/.test(s);
}
