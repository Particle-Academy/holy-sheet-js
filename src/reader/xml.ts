/**
 * Tiny isomorphic XML parser — enough to read OOXML and OpenDocument parts
 * (replacing PHP's SimpleXML). Namespace prefixes are stripped so nodes and
 * attributes are queried by local name, matching how the PHP reader accesses
 * SimpleXML.
 *
 * `parseXml(src, { namespaces: true })` additionally resolves namespaces and
 * keeps mixed content in order, which OpenDocument needs and OOXML does not:
 *
 *   - `ns` on every element, the namespace URI its prefix is bound to;
 *   - `nsAttrs`, attributes keyed `{uri}local` (plain `local` when unprefixed),
 *     because ODS puts `office:value-type` and `calcext:value-type` on the same
 *     cell and local names alone collide;
 *   - `content`, text runs and child elements in document order, because a
 *     paragraph's meaning depends on where its `text:s` falls between the text.
 *
 * Attribute values in that mode get XML attribute-value normalisation (a
 * literal tab, CR or LF is a space), as libxml and expat apply it. Without the
 * option nothing changes, so the xlsx path parses exactly as it always has.
 *
 * This is the one gate for a part of an uploaded file. A part carrying a DOCTYPE
 * is refused, because no real part has one and an internal subset is the entry
 * point for entity expansion. This parser has no entity machinery, so the point
 * here is that it refuses what the other engines refuse and takes no content
 * from inside a subset. Decode through `decodePart`, so the check sees what the
 * part actually says whatever encoding it uses. A test fails if `src/` decodes a
 * part or looks for a DOCTYPE anywhere else. `parseXml` THROWS
 * `UnsupportedFormatException` for a DOCTYPE before the root, as PHP and Python
 * do; malformed XML still reads as null.
 */

import { UnsupportedFormatException } from "../exceptions";

export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
  /** Namespace URI. Only with `{ namespaces: true }`. */
  ns?: string;
  /** Attributes keyed `{uri}local`, or `local` when unprefixed. Only with `{ namespaces: true }`. */
  nsAttrs?: Record<string, string>;
  /** Text runs and child elements in document order. Only with `{ namespaces: true }`. */
  content?: Array<XmlNode | string>;
}

export interface ParseXmlOptions {
  namespaces?: boolean;
}

const XML_NS = "http://www.w3.org/XML/1998/namespace";

function localName(name: string): string {
  const idx = name.indexOf(":");
  return idx >= 0 ? name.slice(idx + 1) : name;
}

function unescapeXml(s: string): string {
  if (s.indexOf("&") < 0) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|\w+);/g, (m, ent: string) => {
    switch (ent) {
      case "amp":
        return "&";
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      case "apos":
        return "'";
      default:
        if (ent[0] === "#") {
          const code = ent[1] === "x" || ent[1] === "X" ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
          return Number.isFinite(code) ? String.fromCodePoint(code) : m;
        }
        return m;
    }
  });
}

function findTagEnd(src: string, from: number): number {
  let quote = "";
  for (let i = from; i < src.length; i++) {
    const ch = src[i]!;
    if (quote) {
      if (ch === quote) quote = "";
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ">") {
      return i;
    }
  }
  return src.length;
}

const ATTR_RE = /([^\s=/]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

interface RawTag {
  name: string;
  attrs: Record<string, string>;
  raw: Array<[string, string]>;
}

function parseTag(inner: string): RawTag {
  const trimmed = inner.trim();
  let i = 0;
  while (i < trimmed.length && !/\s/.test(trimmed[i]!)) i++;
  const name = trimmed.slice(0, i);
  const attrs: Record<string, string> = {};
  const raw: Array<[string, string]> = [];
  const rest = trimmed.slice(i);
  for (const m of rest.matchAll(ATTR_RE)) {
    const rawValue = m[3] ?? m[4] ?? "";
    attrs[localName(m[1]!)] = unescapeXml(rawValue);
    raw.push([m[1]!, rawValue]);
  }
  return { name, attrs, raw };
}

const UTF8 = new TextDecoder();

/**
 * The encoding a part is written in, or null when this port cannot decode it.
 * Reads the byte order first, as an XML parser does, then the declaration.
 */
function partEncoding(b: Uint8Array): string | null {
  if (b[0] === 0xff && b[1] === 0xfe) return b[2] === 0 && b[3] === 0 ? null : "utf-16le";
  if (b[0] === 0xfe && b[1] === 0xff) return "utf-16be";
  if (b[0] === 0 && b[1] === 0) return null;
  if (b[0] === 0x3c && b[1] === 0) return b[2] === 0 ? null : "utf-16le";
  if (b[0] === 0 && b[1] === 0x3c) return "utf-16be";
  const start = b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf ? 3 : 0;
  const head = String.fromCharCode(...b.subarray(start, start + 200));
  const label = /^<\?xml\s[^>]*?encoding\s*=\s*["']([A-Za-z0-9._-]+)["']/.exec(head)?.[1]?.toLowerCase();
  // Bytes that got this far are not UTF-16, whatever the declaration claims.
  if (label === undefined || label === "utf-8" || label.startsWith("utf-16") || label === "utf16") return "utf-8";
  try {
    new TextDecoder(label);
    return label;
  } catch {
    return null;
  }
}

function decodeOrNull(bytes: Uint8Array): string | null {
  const encoding = partEncoding(bytes);
  if (encoding === null) return null;
  return encoding === "utf-8" ? UTF8.decode(bytes) : new TextDecoder(encoding).decode(bytes);
}

/** A part's text, or "" when this port cannot decode it, which parses as nothing. */
export function decodePart(bytes: Uint8Array): string {
  return decodeOrNull(bytes) ?? "";
}

/** Whether a part is refused, as opposed to merely malformed or fine. */
export function partIsRefused(bytes: Uint8Array): boolean {
  const text = decodeOrNull(bytes);
  return text === null || walk(text).refused;
}

/**
 * The part's root, or null when it is empty or malformed.
 *
 * @throws UnsupportedFormatException when a DOCTYPE comes before the root. Read
 *   as an absent part, a worksheet carrying one came back as an EMPTY sheet
 *   (PHP 2.11.5 raises too). Any other declaration is still refused, and still
 *   reads as absent, because PHP's libxml and Python's expat call it malformed.
 */
export function parseXml(src: string, options: ParseXmlOptions = {}): XmlNode | null {
  const walked = walk(src, options);
  if (walked.doctype) throw UnsupportedFormatException.refusedPart();
  return walked.root;
}

function walk(
  src: string,
  options: ParseXmlOptions = {},
): { root: XmlNode | null; refused: boolean; doctype?: boolean } {
  const namespaces = options.namespaces === true;
  let i = 0;
  const n = src.length;
  const root: XmlNode = { name: "#root", attrs: {}, children: [], text: "" };
  if (namespaces) root.content = [];
  const stack: XmlNode[] = [root];
  const scopes: Array<Record<string, string>> = [{ xml: XML_NS }];

  const appendText = (text: string): void => {
    const top = stack[stack.length - 1]!;
    top.text += text;
    if (namespaces && text !== "") top.content!.push(text);
  };

  while (i < n) {
    if (src[i] === "<") {
      if (src.startsWith("<!--", i)) {
        const end = src.indexOf("-->", i);
        i = end < 0 ? n : end + 3;
        continue;
      }
      if (src.startsWith("<![CDATA[", i)) {
        const end = src.indexOf("]]>", i);
        appendText(src.slice(i + 9, end < 0 ? n : end));
        i = end < 0 ? n : end + 3;
        continue;
      }
      if (src.startsWith("<?", i)) {
        const end = src.indexOf("?>", i);
        i = end < 0 ? n : end + 2;
        continue;
      }
      // A declaration: a DOCTYPE, or nothing a part may hold. Refused here rather
      // than skipped, because its end cannot be found without parsing its subset.
      // `doctype` marks the case the other engines raise on: a DOCTYPE in the
      // prolog, before any element.
      if (src.startsWith("<!", i)) {
        const doctype = root.children.length === 0 && src.slice(i, i + 9).toLowerCase() === "<!doctype";
        return { root: null, refused: true, doctype };
      }
      if (src[i + 1] === "/") {
        const end = src.indexOf(">", i);
        if (stack.length > 1) {
          stack.pop();
          if (namespaces) scopes.pop();
        }
        i = end < 0 ? n : end + 1;
        continue;
      }
      const end = findTagEnd(src, i + 1);
      const tagContent = src.slice(i + 1, end);
      const selfClosing = tagContent.endsWith("/");
      const { name, attrs, raw } = parseTag(selfClosing ? tagContent.slice(0, -1) : tagContent);
      const node: XmlNode = { name: localName(name), attrs, children: [], text: "" };
      const parent = stack[stack.length - 1]!;
      parent.children.push(node);

      let scope: Record<string, string> | undefined;
      if (namespaces) {
        scope = resolveNamespaces(node, name, raw, scopes[scopes.length - 1]!);
        parent.content!.push(node);
      }
      if (!selfClosing) {
        stack.push(node);
        if (scope) scopes.push(scope);
      }
      i = end + 1;
    } else {
      const next = src.indexOf("<", i);
      const stop = next < 0 ? n : next;
      appendText(unescapeXml(src.slice(i, stop)));
      i = stop;
    }
  }

  return { root: root.children[0] ?? null, refused: false };
}

function resolveNamespaces(
  node: XmlNode,
  qname: string,
  raw: Array<[string, string]>,
  parentScope: Record<string, string>,
): Record<string, string> {
  let scope = parentScope;
  for (const [key, value] of raw) {
    if (key === "xmlns" || key.startsWith("xmlns:")) {
      if (scope === parentScope) scope = { ...parentScope };
      scope[key === "xmlns" ? "" : key.slice(6)] = unescapeXml(value);
    }
  }

  const colon = qname.indexOf(":");
  node.ns = scope[colon >= 0 ? qname.slice(0, colon) : ""] ?? "";
  node.content = [];

  const nsAttrs: Record<string, string> = {};
  for (const [key, value] of raw) {
    if (key === "xmlns" || key.startsWith("xmlns:")) continue;
    const idx = key.indexOf(":");
    const normalised = unescapeXml(value.replace(/[\t\n\r]/g, " "));
    if (idx < 0) {
      nsAttrs[key] = normalised;
    } else {
      nsAttrs[`{${scope[key.slice(0, idx)] ?? ""}}${key.slice(idx + 1)}`] = normalised;
    }
  }
  node.nsAttrs = nsAttrs;

  return scope;
}

/** First child element with the given local name. */
export function el(node: XmlNode | null | undefined, name: string): XmlNode | undefined {
  return node?.children.find((c) => c.name === name);
}

/** All child elements with the given local name. */
export function els(node: XmlNode | null | undefined, name: string): XmlNode[] {
  return node ? node.children.filter((c) => c.name === name) : [];
}

/** Attribute by local name, or undefined. */
export function at(node: XmlNode | null | undefined, name: string): string | undefined {
  return node?.attrs[name];
}
