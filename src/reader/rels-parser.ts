import { els, parseXml } from "./xml";

export interface Rel {
  Target: string;
  Type: string;
}

/** Parses OOXML *.rels files → rId → {Target, Type}. Mirrors PHP `RelsParser`. */
export const RelsParser = {
  parse(relsXml: string): Record<string, Rel> {
    const xml = parseXml(relsXml);
    if (!xml) return {};
    const rels: Record<string, Rel> = {};
    for (const rel of els(xml, "Relationship")) {
      const id = rel.attrs["Id"] ?? "";
      rels[id] = { Target: rel.attrs["Target"] ?? "", Type: rel.attrs["Type"] ?? "" };
    }
    return rels;
  },

  byType(rels: Record<string, Rel>, typeUriContains: string): Record<string, Rel> {
    const out: Record<string, Rel> = {};
    for (const [id, rel] of Object.entries(rels)) {
      if (rel.Type.includes(typeUriContains)) out[id] = rel;
    }
    return out;
  },
};
