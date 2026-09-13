import { UnsupportedFormatException } from "../exceptions";
import { unzipSync } from "../zip";
import { phpTrim } from "./ods/ns";

export type SpreadsheetFormat = "xlsx" | "ods";

/** The spreadsheet media types, and the template variant (.ots) which is read the same way. */
const ODS_MIMETYPES = new Set([
  "application/vnd.oasis.opendocument.spreadsheet",
  "application/vnd.oasis.opendocument.spreadsheet-template",
]);

const decoder = new TextDecoder();

/**
 * Which reader a file needs, decided from its contents, never its extension.
 * Mirrors PHP `Reader\FormatSniffer`.
 *
 * An OpenDocument package names itself in its `mimetype` entry; an xlsx has no
 * such entry and is recognised by `xl/workbook.xml`. Anything else is refused
 * by name. The unzipped entries are returned so the reader does not unzip again.
 */
export const FormatSniffer = {
  sniff(
    bytes: Uint8Array,
    path: string | null = null,
  ): { format: SpreadsheetFormat; files: Record<string, Uint8Array> } {
    let files: Record<string, Uint8Array>;
    try {
      files = unzipSync(bytes);
    } catch (e) {
      // A real zip this reader cannot inflate is a different failure; say so.
      if (e instanceof Error && e.message.includes("unsupported zip method")) throw e;
      throw UnsupportedFormatException.notAZip(path);
    }

    const mimetype = files["mimetype"] ? phpTrim(decoder.decode(files["mimetype"])) : null;
    if (mimetype !== null && ODS_MIMETYPES.has(mimetype)) return { format: "ods", files };
    if (files["xl/workbook.xml"]) return { format: "xlsx", files };

    throw UnsupportedFormatException.unknownPackage(path, mimetype);
  },
};
