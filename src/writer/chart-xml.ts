import type { ChartRange } from "../workbook/chart-range";
import type { NormalizedChart } from "../workbook/sheet";
import type { Workbook } from "../workbook/workbook";
import { formatFloat } from "../util";

/**
 * The parts a native chart needs. Mirrors PHP `Writer\ChartXml` byte for byte:
 * a drawing per sheet (anchors), a chart part per chart (DrawingML over cell
 * ranges with cached values).
 */

const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const ChartXml = {
  drawing(charts: NormalizedChart[], firstChartNum: number): string {
    let anchors = "";
    charts.forEach((chart, k) => {
      const id = k + 2;
      const num = firstChartNum + k;
      anchors +=
        '<xdr:twoCellAnchor editAs="oneCell">' +
        marker("from", chart.from) +
        marker("to", chart.to) +
        '<xdr:graphicFrame macro="">' +
        `<xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${num}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
        `<a:graphic><a:graphicData uri="${NS_C}">` +
        `<c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="rId${k + 1}"/>` +
        "</a:graphicData></a:graphic>" +
        "</xdr:graphicFrame>" +
        "<xdr:clientData/>" +
        "</xdr:twoCellAnchor>";
    });
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${NS_A}">` +
      anchors +
      "</xdr:wsDr>"
    );
  },

  drawingRels(charts: NormalizedChart[], firstChartNum: number): string {
    let rels = "";
    charts.forEach((_, k) => {
      rels += `<Relationship Id="rId${k + 1}" Type="${NS_R}/chart" Target="../charts/chart${firstChartNum + k}.xml"/>`;
    });
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels +
      "</Relationships>"
    );
  },

  chart(chart: NormalizedChart, workbook: Workbook): string {
    const type = chart.type;
    let series = "";
    chart.series.forEach((s, i) => {
      series += seriesXml(i, s, chart, workbook);
    });
    const labels = chart.dataLabels ? dataLabels(type, chart.format) : "";
    let plot: string;
    if (type === "line") {
      plot = '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series + labels + '<c:marker val="1"/><c:axId val="10"/><c:axId val="20"/></c:lineChart>';
    } else if (type === "pie") {
      plot = '<c:pieChart><c:varyColors val="1"/>' + series + labels + '<c:firstSliceAng val="0"/></c:pieChart>';
    } else {
      plot =
        `<c:barChart><c:barDir val="${type === "bar" ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
        series +
        labels +
        '<c:gapWidth val="80"/><c:axId val="10"/><c:axId val="20"/></c:barChart>';
    }
    const axesXml = type === "pie" ? "" : axes(type === "bar", chart.format);
    const legend = chart.legend ?? (chart.series.length > 1 || type === "pie");
    const title =
      chart.title !== null
        ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr>' +
          `<a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>${text(chart.title)}</a:t></a:r></a:p></c:rich></c:tx>` +
          '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
        : '<c:autoTitleDeleted val="1"/>';

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
      '<c:roundedCorners val="0"/>' +
      "<c:chart>" +
      title +
      "<c:plotArea><c:layout/>" +
      plot +
      axesXml +
      "</c:plotArea>" +
      (legend ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : "") +
      '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>' +
      "</c:chart>" +
      "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>" +
      textProps("595959", 900) +
      "</c:chartSpace>"
    );
  },
};

function marker(tag: string, at: [number, number]): string {
  return `<xdr:${tag}><xdr:col>${at[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${at[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
}

function seriesXml(i: number, s: NormalizedChart["series"][number], chart: NormalizedChart, workbook: Workbook): string {
  const type = chart.type;
  const name = s.name !== null ? `<c:tx><c:v>${text(s.name)}</c:v></c:tx>` : "";
  let fill = "";
  if (s.color !== null && type !== "pie") {
    fill =
      type === "line"
        ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill><a:round/></a:ln></c:spPr>`
        : `<c:spPr><a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill></c:spPr>`;
  }
  const invert = type === "column" || type === "bar" ? '<c:invertIfNegative val="0"/>' : "";
  const mark = type === "line" ? '<c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>' : "";
  const cat = chart.categories !== null ? "<c:cat>" + strRef(chart.categories, workbook) + "</c:cat>" : "";
  const smooth = type === "line" ? '<c:smooth val="0"/>' : "";
  // A coloured point (dPt) follows the series' own fill, invert and marker,
  // and precedes its categories. Mirrors PHP ChartXml::series.
  let points = "";
  for (const p of s.points) {
    const solid = `<a:solidFill><a:srgbClr val="${p.color}"/></a:solidFill>`;
    points +=
      `<c:dPt><c:idx val="${p.index}"/>` +
      (type === "line"
        ? `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${solid}<a:ln>${solid}</a:ln></c:spPr></c:marker><c:bubble3D val="0"/>`
        : type === "pie"
          ? `<c:bubble3D val="0"/><c:spPr>${solid}</c:spPr>`
          : `<c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr>${solid}</c:spPr>`) +
      "</c:dPt>";
  }
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` + name + fill + invert + mark + points + cat + "<c:val>" + numRef(s.values, workbook) + "</c:val>" + smooth + "</c:ser>";
}

function strRef(range: ChartRange, workbook: Workbook): string {
  let pts = "";
  const vals = values(range, workbook);
  vals.forEach((v, idx) => {
    if (v !== null) pts += `<c:pt idx="${idx}"><c:v>${text(category(v))}</c:v></c:pt>`;
  });
  return `<c:strRef><c:f>${text(range.formula())}</c:f><c:strCache><c:ptCount val="${vals.length}"/>` + pts + "</c:strCache></c:strRef>";
}

function numRef(range: ChartRange, workbook: Workbook): string {
  let pts = "";
  const vals = values(range, workbook);
  vals.forEach((v, idx) => {
    if (typeof v === "number") pts += `<c:pt idx="${idx}"><c:v>${formatFloat(v)}</c:v></c:pt>`;
  });
  return (
    `<c:numRef><c:f>${text(range.formula())}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    pts +
    "</c:numCache></c:numRef>"
  );
}

/** Each cell's cached value: a formula cell's cached value, else its value. */
function values(range: ChartRange, workbook: Workbook): (string | number | boolean | null)[] {
  const sheet = workbook.sheets.find((s) => s.name === range.sheet);
  return range.addresses().map((address) => {
    const cell = sheet?.cells.get(address);
    if (cell === undefined) return null;
    const v = cell.formula !== null ? cell.cachedValue : cell.value;
    return v === undefined ? null : (v as string | number | boolean | null);
  });
}

function category(v: string | number | boolean): string {
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return formatFloat(v);
  return v;
}

function dataLabels(type: string, format: string | null): string {
  const pos = type === "line" ? "t" : type === "pie" ? "bestFit" : "outEnd";
  return (
    "<c:dLbls>" +
    (format !== null ? `<c:numFmt formatCode="${attr(format)}" sourceLinked="0"/>` : "") +
    "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>" +
    textProps("404040", 900) +
    `<c:dLblPos val="${pos}"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>` +
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>' +
    "</c:dLbls>"
  );
}

function axes(horizontal: boolean, format: string | null): string {
  const catPos = horizontal ? "l" : "b";
  const valPos = horizontal ? "b" : "l";
  const numFmt = format !== null ? `<c:numFmt formatCode="${attr(format)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>';
  return (
    `<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="${horizontal ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/>` +
    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>' +
    '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>' +
    '<c:crossAx val="20"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
    '<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    `<c:delete val="0"/><c:axPos val="${valPos}"/>` +
    '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E5E7EB"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' +
    numFmt +
    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
    "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>" +
    `<c:crossAx val="10"/><c:crosses val="${horizontal ? "max" : "autoZero"}"/><c:crossBetween val="between"/></c:valAx>`
  );
}

function textProps(color: string, size: number): string {
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

/** Element text: & < > only (PHP htmlspecialchars ENT_XML1 | ENT_NOQUOTES). */
function text(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Attribute value: also both quotes (ENT_XML1 | ENT_QUOTES). */
function attr(s: string): string {
  return text(s).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
