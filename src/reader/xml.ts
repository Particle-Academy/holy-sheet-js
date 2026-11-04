/**
 * Tiny isomorphic XML parser — enough to read OOXML parts (replacing PHP's
 * SimpleXML). Namespace prefixes are stripped so nodes/attributes are queried
 * by local name, matching how the PHP reader accesses SimpleXML.
 *
 * This is the one gate for a part of an uploaded file. A part carrying a DOCTYPE
 * is refused, because no real part has one and an internal subset is the entry
 * point for entity expansion. This parser has no entity machinery, so the point
 * here is that it refuses what the other engines refuse and takes no content
 * from inside a subset. Decode through `decodePart`, so the check sees what the
 * part actually says whatever encoding it uses. A test fails if `src/` decodes a
 * part or looks for a DOCTYPE anywhere else.
 */

export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

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

function parseTag(inner: string): { name: string; attrs: Record<string, string> } {
  const trimmed = inner.trim();
  let i = 0;
  while (i < trimmed.length && !/\s/.test(trimmed[i]!)) i++;
  const name = trimmed.slice(0, i);
  const attrs: Record<string, string> = {};
  const rest = trimmed.slice(i);
  for (const m of rest.matchAll(ATTR_RE)) {
    const key = localName(m[1]!);
    const value = unescapeXml(m[3] ?? m[4] ?? "");
    attrs[key] = value;
  }
  return { name, attrs };
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

/** A part's text, or "" when this port cannot decode it, which parses as nothing. */
export function decodePart(bytes: Uint8Array): string {
  const encoding = partEncoding(bytes);
  if (encoding === null) return "";
  return encoding === "utf-8" ? UTF8.decode(bytes) : new TextDecoder(encoding).decode(bytes);
}

export function parseXml(src: string): XmlNode | null {
  let i = 0;
  const n = src.length;
  const root: XmlNode = { name: "#root", attrs: {}, children: [], text: "" };
  const stack: XmlNode[] = [root];

  while (i < n) {
    if (src[i] === "<") {
      if (src.startsWith("<!--", i)) {
        const end = src.indexOf("-->", i);
        i = end < 0 ? n : end + 3;
        continue;
      }
      if (src.startsWith("<![CDATA[", i)) {
        const end = src.indexOf("]]>", i);
        stack[stack.length - 1]!.text += src.slice(i + 9, end < 0 ? n : end);
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
      if (src.startsWith("<!", i)) return null;
      if (src[i + 1] === "/") {
        const end = src.indexOf(">", i);
        if (stack.length > 1) stack.pop();
        i = end < 0 ? n : end + 1;
        continue;
      }
      const end = findTagEnd(src, i + 1);
      const tagContent = src.slice(i + 1, end);
      const selfClosing = tagContent.endsWith("/");
      const { name, attrs } = parseTag(selfClosing ? tagContent.slice(0, -1) : tagContent);
      const node: XmlNode = { name: localName(name), attrs, children: [], text: "" };
      stack[stack.length - 1]!.children.push(node);
      if (!selfClosing) stack.push(node);
      i = end + 1;
    } else {
      const next = src.indexOf("<", i);
      const stop = next < 0 ? n : next;
      stack[stack.length - 1]!.text += unescapeXml(src.slice(i, stop));
      i = stop;
    }
  }

  return root.children[0] ?? null;
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
