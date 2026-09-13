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
 */

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

export function parseXml(src: string, options: ParseXmlOptions = {}): XmlNode | null {
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
      if (src.startsWith("<!", i)) {
        const end = src.indexOf(">", i);
        i = end < 0 ? n : end + 1;
        continue;
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

  return root.children[0] ?? null;
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
