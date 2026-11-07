import type { ReadLimits } from "./read-limits";
import { el, els, parseXml, type XmlNode } from "./xml";

/** Parses xl/sharedStrings.xml → string[] by index. Mirrors PHP `SharedStringsParser`. */
export const SharedStringsParser = {
  /** The table is counted before any entry is rendered (`maxSharedStrings`). */
  parse(xml: string, limits: ReadLimits | null = null): string[] {
    // Counted while streaming, before the parse builds the table's tree.
    limits?.sharedStringsPart(xml);
    const doc = parseXml(xml);
    if (!doc) return [];
    return els(doc, "si").map(renderSi);
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
