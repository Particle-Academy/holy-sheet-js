import { els, parseXml } from "./xml";

export interface Rel {
  Target: string;
  Type: string;
  TargetMode: "Internal" | "External";
}

/**
 * Parses OOXML *.rels files → rId → {Target, Type, TargetMode}, and resolves
 * their targets. Mirrors PHP `RelsParser`.
 *
 * Targets resolve per OPC (ECMA-376 Part 2): against the directory of the part
 * that OWNS the relationship, unless the target starts with "/", which makes it
 * package-absolute. Excel writes relative targets; openpyxl, LibreOffice and
 * others write absolute ones, and both are normal. An External relationship is
 * a URI, never a part.
 */
export const RelsParser = {
  parse(relsXml: string): Record<string, Rel> {
    const xml = parseXml(relsXml);
    if (!xml) return {};
    const rels: Record<string, Rel> = {};
    for (const rel of els(xml, "Relationship")) {
      const id = rel.attrs["Id"] ?? "";
      rels[id] = {
        Target: rel.attrs["Target"] ?? "",
        Type: rel.attrs["Type"] ?? "",
        TargetMode: rel.attrs["TargetMode"] === "External" ? "External" : "Internal",
      };
    }
    return rels;
  },

  /** The INTERNAL relationships whose Type URI contains the fragment; External ones name no part. */
  byType(rels: Record<string, Rel>, typeUriContains: string): Record<string, Rel> {
    const out: Record<string, Rel> = {};
    for (const [id, rel] of Object.entries(rels)) {
      if (rel.Type.includes(typeUriContains) && !RelsParser.isExternal(rel)) out[id] = rel;
    }
    return out;
  },

  isExternal(rel: Pick<Rel, "TargetMode"> | { TargetMode?: string }): boolean {
    return (rel.TargetMode ?? "Internal") === "External";
  },

  /**
   * The part name `target` names, from the part `source` that owns the
   * relationship ("" for the package root). "." and ".." normalise; ".." never
   * climbs above the package root.
   */
  resolve(source: string, target: string): string {
    const combined = target.startsWith("/")
      ? target
      : (source.includes("/") ? source.slice(0, source.lastIndexOf("/")) : "") + "/" + target;
    const parts: string[] = [];
    for (const segment of combined.split("/")) {
      if (segment === "..") parts.pop();
      else if (segment !== "" && segment !== ".") parts.push(segment);
    }
    return parts.join("/");
  },

  /** xl/worksheets/data.xml → xl/worksheets/_rels/data.xml.rels. */
  relsPathOf(part: string): string {
    const slash = part.lastIndexOf("/");
    return slash < 0 ? `_rels/${part}.rels` : `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
  },
};
