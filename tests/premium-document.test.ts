import { describe, it, expect } from "vitest";
import { Agent, unzipSync } from "../src";

/**
 * Can this package produce a document someone would be PROUD to send?
 *
 * The Node half of `holy-sheet`'s `PremiumDocumentTest`. The other suites prove
 * each feature works ON ITS OWN — a currency column, a merge, a frozen pane.
 * Every one passes, and none asks whether you can turn all of it on at once and
 * get a document with flair rather than a grid of plain text.
 *
 * Styles in xlsx are a DEDUPED TABLE: fonts, fills, borders and number formats
 * are pooled and referenced by index. A per-feature test writes one styled cell,
 * so index 1 is always the thing it just made. Combine eight features and the
 * indices start colliding, and the classic result is a workbook where the last
 * style silently wins and the earlier ones read as unstyled.
 *
 * **A dropped style is invisible.** The file opens, Excel shows no error, and
 * the document is merely plain. Nobody files a bug against a spreadsheet that
 * looks boring — they conclude the library is boring. So these assert the
 * ARTIFACT: unzip, and look for the feature in the XML. "It wrote a file" is a
 * check that passes just as happily for a document with no formatting at all.
 */

const dec = new TextDecoder();

/** Read one part out of a written workbook. */
const partOf = (schema: unknown, name: string): string => {
  const files = unzipSync(Agent.toBytes(schema as never));
  const part = files[name];
  // A missing part and an empty one are different failures, and the difference
  // is the whole diagnosis: absent means the writer never emitted it.
  expect(part, `workbook has no ${name}`).toBeDefined();
  return dec.decode(part!);
};

const countOf = (haystack: string, needle: string): number => haystack.split(needle).length - 1;

/** Everything a rich workbook uses, on ONE sheet. */
const premiumWorkbook = () => ({
  sheets: [
    {
      name: "Q3 Revenue",
      theme: "default",
      columns: [
        { header: "Region", width: 24 },
        { header: "Revenue", type: "currency", currency: "USD", decimals: 2 },
        { header: "Growth", type: "percent", decimals: 1 },
        { header: "Closed", type: "date" },
      ],
      rows: [
        ["North", 1250000.5, 0.184, "2026-09-30"],
        ["South", 980400.25, -0.042, "2026-09-30"],
        ["EMEA", 2100000.0, 0.311, "2026-09-30"],
      ],
      totals: { Revenue: "sum", Growth: "avg" },
      mergedRegions: [{ start: "A7", end: "D7" }],
      frozenRows: 1,
      frozenCols: 1,
      columnWidths: { 0: 200, 1: 150 },
      cells: {
        // A title band beneath the table: large, bold, reversed out of a dark
        // fill, merged across all four columns.
        A7: {
          value: "Q3 REVENUE BY REGION",
          format: {
            bold: true,
            fontSize: 20,
            color: "#FFFFFF",
            backgroundColor: "#1F3864",
            textAlign: "center",
          },
        },
        // A callout: italic, coloured, ruled above and below.
        A9: {
          value: "EMEA led on growth",
          format: {
            italic: true,
            fontSize: 11,
            color: "#C00000",
            borderTop: "#C00000",
            borderBottom: "#C00000",
          },
        },
        B9: { value: "see note", comment: "Reviewed by finance 2026-10-01" },
        // Reaches INTO the table: emphasise one figure without restating the
        // column's currency format. See the inheritance test.
        B4: { value: 2100000.0, format: { bold: true } },
      },
    },
  ],
});

describe("a rich workbook keeps EVERY styling feature, together", () => {
  it("keeps all eight per-cell format fields in one document", () => {
    // The composition check. Each of these has a home in the deduped style
    // table, and they are asserted in one workbook precisely because the
    // per-feature tests cannot see an index collision between them.
    const styles = partOf(premiumWorkbook(), "xl/styles.xml");

    const expected: Record<string, string> = {
      bold: "<b/>",
      italic: "<i/>",
      "font size 20": 'val="20"',
      "font size 11": 'val="11"',
      "white text": "FFFFFFFF",
      "dark navy fill": "FF1F3864",
      "red text": "FFC00000",
      centred: "center",
    };

    const missing = Object.entries(expected)
      .filter(([, needle]) => !styles.includes(needle))
      .map(([label]) => label);

    expect(
      missing,
      `these styles were accepted and never reached styles.xml: ${missing.join(", ")}`,
    ).toEqual([]);
  });

  it("rules the callout above AND below, not just one edge", () => {
    // Borders are the easiest thing to half-implement: an edge that is written
    // for `top` and dropped for `bottom` still produces a document with borders
    // in it, so a loose assertion passes.
    const styles = partOf(premiumWorkbook(), "xl/styles.xml");

    expect(styles, "no top border element").toContain("<top");
    expect(styles, "no bottom border element").toContain("<bottom");
    expect(countOf(styles, "FFC00000")).toBeGreaterThanOrEqual(2);
  });

  it("lays the sheet out: merge, freeze and widths survive together", () => {
    const sheet = partOf(premiumWorkbook(), "xl/worksheets/sheet1.xml");

    expect(sheet, "title band did not merge").toContain('<mergeCell ref="A7:D7"/>');
    expect(sheet, "panes not frozen").toContain('state="frozen"');
    expect(sheet, "column widths not applied").toContain('<col min="1" max="1"');
  });

  it("formats numbers as money, percent and dates rather than raw digits", () => {
    // The difference between a premium document and a boring one is often
    // exactly this: 1250000.5 versus $1,250,000.50. A number format that
    // silently does not apply leaves a correct spreadsheet nobody wants.
    const styles = partOf(premiumWorkbook(), "xl/styles.xml");

    expect(styles, "no custom number formats at all").toContain("<numFmts");
    expect(styles, "currency column produced no currency format").toContain("$");
    expect(styles, "percent column produced no percent format").toContain("%");
  });

  it("carries the reviewer comment as a real comment part", () => {
    expect(partOf(premiumWorkbook(), "xl/comments1.xml")).toContain("Reviewed by finance");
  });

  it("writes totals as FORMULAS, so the document recalculates", () => {
    // A totals row of baked numbers looks identical and is dead on arrival:
    // edit a cell above it and the total lies. This is the difference between a
    // report and a picture of a report.
    const sheet = partOf(premiumWorkbook(), "xl/worksheets/sheet1.xml");

    expect(sheet, "totals row contains no formulas").toContain("<f>");
    expect(sheet, "no SUM in the totals row").toContain("SUM(");
  });
});

describe("a sheet keeps BOTH its table and its explicit cells", () => {
  // ## The defect these pin
  //
  // `normalizeSheet` returned the moment `cells` was set, which threw away
  // `columns`, `rows`, `totals` and `theme`. A four-row table with one styled
  // title collapsed to a one-cell workbook.
  //
  // Nothing reported it. `validate()` returned no errors, because the validator
  // says in as many words that a sheet may carry "columns + rows OR cells
  // (sparse map) OR both". The validator permitted the combination and the
  // writer silently dropped half of it — so the file opened cleanly, and was
  // simply missing the report.
  //
  // The PHP twin had the identical bug in the identical place, which is the
  // reason a parity suite never caught it: parity detects DISAGREEMENT, and
  // these two agreed.

  it("does not discard the table when a styled cell is added beside it", () => {
    const sheet = partOf(premiumWorkbook(), "xl/worksheets/sheet1.xml");

    for (const needle of ["Region", "North", "South", "EMEA", "Q3 REVENUE BY REGION"]) {
      expect(sheet, `\`${needle}\` was dropped from the composed sheet`).toContain(needle);
    }

    // Rows 1-5 are the table, 7 and 9 the styled cells. Counting them is what
    // catches the collapse: the old code wrote exactly the `cells`.
    expect(countOf(sheet, "<row ")).toBeGreaterThanOrEqual(7);
  });

  it("keeps the totals FORMULAS when explicit cells are present", () => {
    // The formulas are generated by the same branch the early return skipped,
    // so they went with the table.
    const sheet = partOf(premiumWorkbook(), "xl/worksheets/sheet1.xml");

    expect(sheet, "the totals SUM did not survive").toContain("SUM(B2:B4)");
    expect(sheet, "the totals AVERAGE did not survive").toContain("AVERAGE(C2:C4)");
  });

  it("lets an explicit cell INHERIT the column format it sits in", () => {
    // B4 asks only for `bold`. It must keep the currency format from its
    // column — otherwise emphasising one figure silently turns it back into a
    // raw number, and "2100000" in a column of dollars is a worse document than
    // the one before the change.
    //
    // The style is nowhere near the cell in the file: `s="N"` indexes cellXfs,
    // whose entry names a font and a number format by index in turn. Following
    // that chain is the only way to assert what a cell LOOKS like, which is
    // exactly why a dropped format is so easy to miss.
    const sheet = partOf(premiumWorkbook(), "xl/worksheets/sheet1.xml");
    const styles = partOf(premiumWorkbook(), "xl/styles.xml");

    const styleIndex = /<c r="B4"[^>]*?\ss="(\d+)"/.exec(sheet)?.[1];
    expect(styleIndex, "B4 was written with no style at all").toBeDefined();

    const cellXfs = styles.slice(styles.indexOf("<cellXfs"), styles.indexOf("</cellXfs>"));
    const xf = [...cellXfs.matchAll(/<xf\b[^>]*/g)].map((m) => m[0])[Number(styleIndex)];
    expect(xf, `no cellXfs entry at index ${styleIndex}`).toBeDefined();

    const numFmtId = /numFmtId="(\d+)"/.exec(xf!)?.[1];
    const formatCode = new RegExp(`<numFmt numFmtId="${numFmtId}" formatCode="([^"]*)"`).exec(
      styles,
    )?.[1];
    expect(
      formatCode ?? "",
      `the column currency format was lost under the overlay: got \`${formatCode}\``,
    ).toContain("$");

    const fontId = Number(/fontId="(\d+)"/.exec(xf!)?.[1]);
    const fonts = styles.slice(styles.indexOf("<fonts"), styles.indexOf("</fonts>"));
    const font = [...fonts.matchAll(/<font>[\s\S]*?<\/font>/g)].map((m) => m[0])[fontId];
    expect(font ?? "", "the explicit bold was not applied").toContain("<b/>");
  });

  it("lets an explicit cell OVERRIDE the format beneath it", () => {
    // The other direction, or "inherit" would just mean "ignore".
    const styles = partOf(
      {
        sheets: [
          {
            name: "Override",
            columns: [{ header: "Amount", type: "currency", currency: "USD" }],
            rows: [[42.0]],
            cells: { A2: { value: 42.0, format: { backgroundColor: "#FFFF00" } } },
          },
        ],
      },
      "xl/styles.xml",
    );

    expect(styles, "the overlay fill never reached the style table").toContain("FFFFFF00");
  });
});

describe("cells written the obvious way are not dropped", () => {
  // These are PORT gaps: PHP has always handled them, Node never did. A parity
  // suite cannot see a feature that only one side was ever asked about, so they
  // survived every cross-runtime comparison in the repo.

  it("accepts a BARE SCALAR cell, not only a {value: ...} object", () => {
    // `{A1: 42}` read `cellData.value` off a number, got undefined, and wrote an
    // empty cell. No error — just a hole where the number should be.
    const sheet = partOf(
      { sheets: [{ name: "Bare", cells: { A1: 42, A2: "plain text" } }] },
      "xl/worksheets/sheet1.xml",
    );

    expect(sheet, "a bare numeric cell wrote nothing").toContain("42");
    expect(sheet, "a bare string cell wrote nothing").toContain("plain text");
  });

  it("promotes a bare '=' string to a real formula, in a CELL", () => {
    const sheet = partOf(
      { sheets: [{ name: "F", cells: { A1: 2, A2: "=SUM(A1:A1)" } }] },
      "xl/worksheets/sheet1.xml",
    );

    expect(sheet).toContain("<f>SUM(A1:A1)</f>");
  });

  it("promotes a bare '=' string to a real formula, in a ROW", () => {
    // The row path had the same gap, so one schema produced a formula in PHP
    // and the literal text "=SUM(...)" in Node.
    const sheet = partOf(
      { sheets: [{ name: "F", columns: [{ header: "N" }], rows: [[2], ["=SUM(A2:A2)"]] }] },
      "xl/worksheets/sheet1.xml",
    );

    expect(sheet).toContain("<f>SUM(A2:A2)</f>");
  });

  it("accepts a comment given as a plain string", () => {
    // `comment` was read only as an object, so the string form was accepted by
    // the validator, normalized to null, and lost. No error, no comment.
    const comments = partOf(
      {
        sheets: [
          { name: "Noted", cells: { A1: { value: "x", comment: "a bare string comment" } } },
        ],
      },
      "xl/comments1.xml",
    );

    expect(comments).toContain("a bare string comment");
  });
});

describe("the guard against a style that is accepted and dropped", () => {
  it("proves the assertions can FAIL — an unstyled workbook has none of it", () => {
    // Without this, every assertion above could be passing on boilerplate that
    // appears in any workbook, and the suite would be green for a document with
    // no formatting whatsoever. This is the control.
    const plain = { sheets: [{ name: "Plain", columns: [{ header: "A" }], rows: [["just text"]] }] };

    const styles = partOf(plain, "xl/styles.xml");
    const sheet = partOf(plain, "xl/worksheets/sheet1.xml");

    expect(styles, "a plain workbook somehow contains the premium fill").not.toContain("FF1F3864");
    expect(styles, "a plain workbook somehow contains the title font size").not.toContain(
      'val="20"',
    );
    expect(sheet, "a plain workbook somehow contains a merge").not.toContain("<mergeCell");
  });
});
