import { CellComment } from "../workbook/cell-comment";
import { el, els, parseXml, type XmlNode } from "./xml";

/** Parses xl/commentsN.xml → address → CellComment. Mirrors PHP `CommentsParser`. */
export const CommentsParser = {
  parse(commentsXml: string): Record<string, CellComment> {
    const xml = parseXml(commentsXml);
    if (!xml) return {};

    const authors: string[] = els(el(xml, "authors"), "author").map((a) => a.text);

    const out: Record<string, CellComment> = {};
    for (const c of els(el(xml, "commentList"), "comment")) {
      const ref = c.attrs["ref"] ?? "";
      const authorIdx = c.attrs["authorId"] !== undefined ? parseInt(c.attrs["authorId"], 10) : -1;
      const author = authors[authorIdx] ?? null;
      out[ref] = new CellComment(extractText(el(c, "text")), author);
    }
    return out;
  },
};

function extractText(textNode: XmlNode | undefined): string {
  if (!textNode) return "";
  const parts: string[] = [];
  for (const r of els(textNode, "r")) parts.push(el(r, "t")?.text ?? "");
  for (const t of els(textNode, "t")) parts.push(t.text);
  return parts.join("");
}
