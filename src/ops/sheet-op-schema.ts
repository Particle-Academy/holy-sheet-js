/**
 * JSON Schema for one sheet op — validate ops on the wire, or register the op
 * vocabulary as an LLM tool. Mirrors PHP `HolySheet\Ops\SheetOpSchema`, key for
 * key (pinned byte-for-byte against the PHP package by
 * `tests/sheet-ops-parity.test.ts`).
 *
 * `set_cell`, `set_range` and `set_workbook` are fancy-sheets' `SheetOp`
 * variants, same `type` and same fields, so a stored op stream can drive a live
 * `useSheetSync` session. One difference: `set_workbook.data` here is a Holy
 * Sheet schema, not fancy-sheets' `WorkbookData`.
 */
export const SheetOpSchema = {
  /** Every op type, in the order the variants are listed. */
  TYPES: [
    "set_cell",
    "set_range",
    "set_workbook",
    "clear_cell",
    "insert_rows",
    "delete_rows",
    "insert_columns",
    "delete_columns",
    "add_sheet",
    "remove_sheet",
    "rename_sheet",
    "move_sheet",
    "replace_sheet",
    "set_merged_regions",
    "set_column_widths",
    "set_frozen",
    "set_meta",
  ] as const,

  /** A fresh object on every call, so a caller editing it changes nothing here. */
  jsonSchema(): Record<string, unknown> {
    const address = () => ({ type: "string", pattern: "^[A-Za-z]+[0-9]+$" });
    const sheet = () => ({ type: "string", minLength: 1 });
    const count = () => ({ type: "integer", minimum: 1 });
    const position = () => ({ type: "integer", minimum: 1 });
    const object = () => ({ type: "object" });
    const nullableObject = () => ({ type: ["object", "null"] });
    const value = () => ({ type: ["string", "number", "boolean", "null"] });

    const variants = [
      variant(
        "set_cell",
        { sheet: sheet(), address: address(), value: value(), formula: { type: "string" }, computedValue: value(), format: nullableObject(), comment: nullableObject() },
        ["sheet", "address"],
        "Write one cell. Omitting formula or computedValue clears it; omitting format or comment keeps it; null clears it.",
      ),
      variant(
        "set_range",
        { sheet: sheet(), start: address(), end: address(), values: { type: "array", items: { type: "array", items: value() } } },
        ["sheet", "start", "values"],
        "Write a block of values row-major from start, each as a set_cell with no formula.",
      ),
      variant("set_workbook", { data: object() }, ["data"], "Replace the whole workbook schema."),
      variant("clear_cell", { sheet: sheet(), address: address() }, ["sheet", "address"], "Remove one cell."),
      variant("insert_rows", { sheet: sheet(), at: position(), count: count() }, ["sheet", "at", "count"], "Insert rows before 1-based row `at`; cells, merges below move down."),
      variant("delete_rows", { sheet: sheet(), at: position(), count: count() }, ["sheet", "at", "count"], "Delete rows starting at 1-based row `at`; cells, merges below move up."),
      variant("insert_columns", { sheet: sheet(), at: position(), count: count() }, ["sheet", "at", "count"], "Insert columns before 1-based column `at` (A = 1); cells, merges and widths move right."),
      variant("delete_columns", { sheet: sheet(), at: position(), count: count() }, ["sheet", "at", "count"], "Delete columns starting at 1-based column `at`; cells, merges and widths move left."),
      variant("add_sheet", { index: { type: "integer", minimum: 0 }, sheet: object() }, ["index", "sheet"], "Insert a sheet at a 0-based position."),
      variant("remove_sheet", { sheet: sheet() }, ["sheet"], "Remove a sheet by name."),
      variant("rename_sheet", { sheet: sheet(), name: sheet() }, ["sheet", "name"], "Rename a sheet."),
      variant("move_sheet", { sheet: sheet(), toIndex: { type: "integer", minimum: 0 } }, ["sheet", "toIndex"], "Move a sheet to a 0-based position."),
      variant("replace_sheet", { sheet: sheet(), data: object() }, ["sheet", "data"], "Replace one sheet whole."),
      variant(
        "set_merged_regions",
        { sheet: sheet(), mergedRegions: { type: "array", items: { type: "object", required: ["start", "end"], properties: { start: address(), end: address() } } } },
        ["sheet", "mergedRegions"],
        "Set every merged region of a sheet.",
      ),
      // A LIST too, as PHP 2.3.2 declares. PHP encodes a map whose keys run
      // 0..n-1 as a JSON list, so widths for columns A, B and C arrive as
      // `[120, 80, 140]` and no widths as `[]`; this port passes such a list
      // through a diff unchanged. A list's position is the column index.
      variant(
        "set_column_widths",
        {
          sheet: sheet(),
          columnWidths: {
            type: ["object", "array"],
            items: { type: "number", minimum: 0 },
            additionalProperties: { type: "number", minimum: 0 },
          },
        },
        ["sheet", "columnWidths"],
        "Set every column width of a sheet (0-based column index to pixels; a list is indexed by position); empty removes them.",
      ),
      variant(
        "set_frozen",
        { sheet: sheet(), rows: { type: "integer", minimum: 0 }, cols: { type: "integer", minimum: 0 } },
        ["sheet", "rows", "cols"],
        "Set frozen rows and columns.",
      ),
      variant("set_meta", { meta: nullableObject() }, ["meta"], "Replace the workbook meta, or remove it with null."),
    ];

    return {
      $schema: "http://json-schema.org/draft-07/schema#",
      title: "Holy Sheet op",
      description: "One op from Agent::diff, applied by Agent::reduce.",
      oneOf: variants,
    };
  },
};

function variant(
  type: string,
  properties: Record<string, unknown>,
  required: string[],
  description: string,
): Record<string, unknown> {
  return {
    type: "object",
    description,
    required: ["type", ...required],
    additionalProperties: false,
    properties: { type: { const: type }, ...properties },
  };
}
