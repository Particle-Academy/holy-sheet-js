import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { zipSync } from "../src";

/**
 * The ODS fixtures live in the PHP reference repo, not here, so all three
 * engines read the same bytes (see holy-sheet/tests/fixtures/ods/README.md).
 *
 * Located the way scripts/php-describe.php locates the PHP sources:
 * HOLY_SHEET_PHP_SRC first (CI), the sibling checkout in the .agi envelope
 * second. A missing directory is an error, never a skip: a suite that goes
 * green without its fixtures asserts nothing.
 */
export const ODS_FIXTURES = process.env.HOLY_SHEET_PHP_SRC
  ? join(process.env.HOLY_SHEET_PHP_SRC, "..", "tests", "fixtures", "ods")
  : join(__dirname, "..", "..", "holy-sheet", "tests", "fixtures", "ods");

export function odsFixturePath(name: string): string {
  const path = join(ODS_FIXTURES, name);
  if (!existsSync(path)) {
    throw new Error(
      `ODS fixture not found: ${path}. The fixtures live in the PHP holy-sheet repo; ` +
        "check it out beside this one or set HOLY_SHEET_PHP_SRC to its src/ directory.",
    );
  }
  return path;
}

export function odsFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(odsFixturePath(name)));
}

const NAMESPACES =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"';

/** A minimal .ods from table XML, for constructs no fixture carries. Mirrors the PHP test's ods_package(). */
export function odsPackage(
  tables: string,
  automaticStyles = "",
  mimetype = "application/vnd.oasis.opendocument.spreadsheet",
): Uint8Array {
  const enc = new TextEncoder();
  const content =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    `<office:document-content ${NAMESPACES} office:version="1.3">` +
    `<office:automatic-styles>${automaticStyles}</office:automatic-styles>` +
    `<office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body>` +
    "</office:document-content>";
  return zipSync([
    { name: "mimetype", data: enc.encode(mimetype) },
    { name: "content.xml", data: enc.encode(content) },
  ]);
}
