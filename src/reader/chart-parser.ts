import { ChartRange } from "../workbook/chart-range";
import type { NormalizedChart } from "../workbook/sheet";
import { at, el, els, parseXml, type XmlNode } from "./xml";

/**
 * Charts back into the normalised shape. Mirrors PHP `Reader\ChartParser`
 * lookup for lookup; the PHP side is namespace-aware, this matches local
 * names, which agree on everything this writer and Excel emit.
 */

/** Descendants (not self) with this local name, in document order. */
function descendants(node: XmlNode | undefined, name: string): XmlNode[] {
  const out: XmlNode[] = [];
  const stack: XmlNode[] = node ? [...node.children].reverse() : [];
  while (stack.length > 0) {
    const n = stack.pop()!;
    if (n.name === name) out.push(n);
    for (let i = n.children.length - 1; i >= 0; i--) stack.push(n.children[i]!);
  }
  return out;
}

function load(xml: string): XmlNode | null {
  if (/<!DOCTYPE/i.test(xml)) return null;
  return parseXml(xml);
}

/** Root wrapped so a descendant search also finds the root element itself. */
function wrap(root: XmlNode): XmlNode {
  return { name: "#document", attrs: {}, children: [root], text: "" };
}

export const ChartParser = {
  anchors(drawingXml: string): { from: [number, number]; to: [number, number]; rel: string }[] {
    const root = load(drawingXml);
    if (!root) return [];
    const out: { from: [number, number]; to: [number, number]; rel: string }[] = [];
    for (const anchor of descendants(wrap(root), "twoCellAnchor")) {
      const from = marker(anchor, "from");
      const to = marker(anchor, "to");
      const chart = descendants(anchor, "chart")[0];
      const rel = chart ? (at(chart, "id") ?? "") : "";
      if (from !== null && to !== null && rel !== "") out.push({ from, to, rel });
    }
    return out;
  },

  chart(chartXml: string, sheetName: string, from: [number, number], to: [number, number]): NormalizedChart | null {
    const root = load(chartXml);
    if (!root) return null;
    const doc = wrap(root);
    let plot: XmlNode | undefined;
    let type: NormalizedChart["type"] | null = null;
    for (const tag of ["barChart", "lineChart", "pieChart"]) {
      const found = descendants(doc, tag)[0];
      if (found) {
        plot = found;
        type = tag === "lineChart" ? "line" : tag === "pieChart" ? "pie" : at(el(found, "barDir"), "val") === "bar" ? "bar" : "column";
        break;
      }
    }
    if (!plot || type === null) return null;

    let categories: ChartRange | null = null;
    const series: NormalizedChart["series"] = [];
    for (const ser of els(plot, "ser")) {
      const valRef = formula(ser, "val");
      const values = valRef !== null ? ChartRange.parse(valRef, sheetName) : null;
      if (values === null) continue;
      if (categories === null) {
        const catRef = formula(ser, "cat");
        categories = catRef !== null ? ChartRange.parse(catRef, sheetName) : null;
      }
      series.push({ name: seriesName(ser), values, color: seriesColor(ser), points: points(ser) });
    }
    if (series.length === 0) return null;

    const labels = el(plot, "dLbls");
    const dataLabels = labels !== undefined && at(el(labels, "showVal"), "val") === "1";
    let format: string | null = null;
    for (const host of [labels, descendants(doc, "valAx")[0]]) {
      const numFmt = host ? el(host, "numFmt") : undefined;
      if (numFmt && at(numFmt, "sourceLinked") === "0" && (at(numFmt, "formatCode") ?? "") !== "") {
        format = at(numFmt, "formatCode")!;
        break;
      }
    }

    let title: string | null = null;
    const titleEl = descendants(doc, "title")[0];
    if (titleEl) {
      const text = descendants(titleEl, "t").map((t) => t.text).join("");
      title = text !== "" ? text : null;
    }

    return {
      type,
      title,
      categories,
      series,
      from,
      to,
      legend: descendants(doc, "legend").length > 0 ? true : null,
      dataLabels,
      format,
    };
  },
};

function marker(anchor: XmlNode, tag: string): [number, number] | null {
  const m = el(anchor, tag);
  if (!m) return null;
  const col = descendants(m, "col")[0]?.text.trim() ?? "";
  const row = descendants(m, "row")[0]?.text.trim() ?? "";
  return /^\d+$/.test(col) && /^\d+$/.test(row) ? [parseInt(col, 10), parseInt(row, 10)] : null;
}

function formula(ser: XmlNode, name: string): string | null {
  const host = el(ser, name);
  if (!host) return null;
  for (const ref of ["numRef", "strRef"]) {
    const f = el(el(host, ref), "f");
    if (f && f.text.trim() !== "") return f.text.trim();
  }
  return null;
}

function seriesName(ser: XmlNode): string | null {
  const tx = el(ser, "tx");
  if (!tx) return null;
  const v = el(tx, "v");
  if (v) return v.text;
  const cached = descendants(tx, "v")[0];
  return cached ? cached.text : null;
}

/** Coloured points (c:dPt), by index. Mirrors PHP ChartParser::points. */
function points(ser: XmlNode): { index: number; color: string }[] {
  const byIndex = new Map<number, string>();
  for (const dPt of els(ser, "dPt")) {
    const idx = at(el(dPt, "idx"), "val");
    const clr = descendants(dPt, "srgbClr")[0];
    const color = (clr ? (at(clr, "val") ?? "") : "").toUpperCase();
    if (idx !== undefined && /^[0-9]+$/.test(idx) && /^[0-9A-F]{6}$/.test(color)) {
      byIndex.set(parseInt(idx, 10), color);
    }
  }
  return [...byIndex.entries()].sort((a, b) => a[0] - b[0]).map(([index, color]) => ({ index, color }));
}

function seriesColor(ser: XmlNode): string | null {
  const spPr = el(ser, "spPr");
  if (!spPr) return null;
  const clr = descendants(spPr, "srgbClr")[0];
  const val = (clr ? (at(clr, "val") ?? "") : "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(val) ? val : null;
}
