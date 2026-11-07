import type { ReadLimits } from "./read-limits";
import { el, els, parseXml, type XmlNode } from "./xml";

/** Parses xl/sharedStrings.xml → string[] by index. Mirrors PHP `SharedStringsParser`. */
export const SharedStringsParser = {
  /** The table is counted before any entry is rendered (`maxSharedStrings`). */
  parse(xml: string, limits: ReadLimits | null = null): string[] {
    const doc = parseXml(xml);
    if (!doc) return [];
    const items = els(doc, "si");
    limits?.sharedStrings(items.length);
    return items.map(renderSi);
  },
};

/**
 * One string item's text: its `<t>`, or its `<r><t>` runs joined. Also how a
 * cell's inline `<is>` is read, which has the same content model.
 */
export function renderSi(si: XmlNode): string {
  const t = el(si, "t");
  if (t) return t.text;
  return els(si, "r")
    .map((run) => el(run, "t")?.text ?? "")
    .join("");
}
