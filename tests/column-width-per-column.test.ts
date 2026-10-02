/**
 * A column's OWN `width`, which this writer documented and ignored.
 *
 * `skills/holy-sheet.schema.json` — the tool definition handed to an LLM, shared
 * byte-identically with PHP and Python under a checksum test — describes
 * `columns[].width` as *"Column width in pixels. Same as columnWidths but
 * per-column."* Only the sheet-level `columnWidths` map was ever read, so a width
 * written the way the schema documents it emitted **no `<cols>` element at all**:
 * no error, no repair note, nothing from `validateAndRepair` at any layer.
 *
 * Reported against PHP as holy-sheet#8 by MOIC, whose owner's complaint was
 * "can't style spreadsheets at all". Measured, that was mostly this: a real
 * account name truncates and a correctly-formatted currency renders as `#####`,
 * which is what a reader actually sees. Fixed in PHP 2.4.0; this is the twin.
 *
 * PRECEDENCE, matching PHP exactly: the sheet-level MAP IS APPLIED LAST AND
 * WINS. It is the mechanism that already worked, so anyone who moved to it to
 * route around this bug must not then find a leftover `width` quietly overriding
 * them.
 */
import { describe, expect, it } from "vitest";
import { Agent } from "../src";
import { Normalizer } from "../src/schema/normalizer";

const widthsOf = (doc: unknown): [number, number][] => {
  const workbook = new Normalizer().normalize(doc as never);
  const map = (workbook.sheets[0] as unknown as { columnWidths: Map<number, number> }).columnWidths;
  return [...map.entries()];
};

const withColumns = (columns: unknown[], extra: Record<string, unknown> = {}) => ({
  sheets: [{ name: "S", columns, rows: [["a", "b", "c"]], ...extra }],
});

describe("columns[].width", () => {
  it("is honoured, which is the whole bug", () => {
    expect(
      widthsOf(
        withColumns([
          { header: "Account", width: 260 },
          { header: "Amount", width: 120 },
        ]),
      ),
    ).toEqual([
      [0, 260],
      [1, 120],
    ]);
  });

  it("widens the column at its POSITION, so a column without a width is skipped", () => {
    expect(
      widthsOf(
        withColumns([
          { header: "A" },
          { header: "B", width: 90 },
          { header: "C" },
        ]),
      ),
    ).toEqual([[1, 90]]);
  });

  it("lets the sheet-level map WIN on a conflict", () => {
    // Deliberate, and the reason is migration: `columnWidths` is what worked, so
    // a consumer who moved to it must not be overridden by a stale `width`.
    expect(
      widthsOf(
        withColumns([{ header: "Account", width: 260 }], { columnWidths: { 0: 400 } }),
      ),
    ).toEqual([[0, 400]]);
  });

  it("merges the two sources rather than letting one replace the other", () => {
    expect(
      widthsOf(
        withColumns(
          [
            { header: "A", width: 100 },
            { header: "B", width: 200 },
          ],
          { columnWidths: { 2: 300 } },
        ),
      ),
    ).toEqual([
      [0, 100],
      [1, 200],
      [2, 300],
    ]);
  });

  it("emits <col> elements in ASCENDING column order", () => {
    // Merging two sources means insertion order is no longer column order, and a
    // `<cols>` whose children are out of order is not what the format expects.
    const entries = widthsOf(
      withColumns([{ header: "A" }, { header: "B" }, { header: "C", width: 70 }], {
        columnWidths: { 0: 50 },
      }),
    );

    expect(entries.map(([index]) => index)).toEqual([0, 2]);
  });

  it("applies the same rule to a width as the sheet-level map does", () => {
    // One rule across both sources, or `columns[].width` becomes a second place
    // where "abc" means column A.
    expect(widthsOf(withColumns([{ header: "A", width: "120" }]))).toEqual([[0, 120]]);
    expect(widthsOf(withColumns([{ header: "A", width: -5 }]))).toEqual([]);
    expect(widthsOf(withColumns([{ header: "A", width: "wide" }]))).toEqual([]);
  });

  it("reports an invalid width by its path instead of dropping it in silence", () => {
    // The original defect's real cost was silence: an agent composes against the
    // schema doc, gets no error, and never learns the field was discarded.
    const errors = Agent.validate(withColumns([{ header: "A", width: "wide" }]) as never);

    expect(errors.length).toBeGreaterThan(0);
    // Assert on the PATH, not a stringified error. The point of the error is
    // that it names which column's width was rejected.
    expect(errors.map((e) => e.path)).toContain("sheets[0].columns[0].width");
  });

  it("accepts a string column definition, which carries no width", () => {
    // `columns: ["Account", "Amount"]` is a documented shorthand and must not
    // throw when the merge looks for `.width` on a string.
    expect(widthsOf(withColumns(["Account", "Amount"]))).toEqual([]);
  });
});
