import { el, els, parseXml, type XmlNode } from "./xml";

/** Parses xl/sharedStrings.xml → string[] by index. Mirrors PHP `SharedStringsParser`. */
export const SharedStringsParser = {
  parse(xml: string): string[] {
    const doc = parseXml(xml);
    if (!doc) return [];
    return els(doc, "si").map(renderSi);
  },
};

function renderSi(si: XmlNode): string {
  const t = el(si, "t");
  if (t) return t.text;
  return els(si, "r")
    .map((run) => el(run, "t")?.text ?? "")
    .join("");
}
