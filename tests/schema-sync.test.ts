import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * The JSON Schema is duplicated in the PHP repo and kept in sync BY HAND.
 *
 * `holy-sheet/skills/holy-sheet.schema.json` and
 * `holy-sheet-js/src/holy-sheet.schema.json` are byte-identical copies of one
 * contract, maintained by remembering to edit both. Nothing checked that, and
 * this is the file handed to an LLM as the tool definition — so a one-sided
 * edit does not fail a build, it changes what an agent is told the API is on
 * one backend and not the other.
 *
 * The shared conformance policy names this case directly:
 *
 *   > Contracts that are byte-identical copies get a checksum test. A five-line
 *   > "this file hashes to X" test in each repo makes an unsynchronised edit a
 *   > build failure instead of a discovery.
 *
 * ## How to change the schema
 *
 * Edit BOTH copies, run either suite, and paste the new hash into BOTH tests.
 * That is deliberately mildly annoying: the annoyance is the mechanism. Update
 * one side and the other repo goes red on its next push, which is the whole
 * point — the alternative is silence.
 *
 * The twin lives at `holy-sheet/tests/Unit/SchemaSyncTest.php` and pins the
 * same constant.
 *
 * ## Why the content is normalised before hashing
 *
 * Neither repo has a `.gitattributes`, so the checkout decides line endings.
 * This file is stored LF in git and lands CRLF on a Windows working tree — 196
 * of them. A checksum over the RAW bytes would therefore pass on Linux CI and
 * fail on the maintainer's own machine, for a reason that has nothing to do
 * with the two copies being out of sync. That is a worse failure than the one
 * being prevented, because it trains people to distrust the test.
 *
 * So `\r` is stripped first. Every real content change still moves the hash;
 * only the line-ending representation is forgiven.
 */
const SHARED_SCHEMA_SHA256 =
  "01d701da2cc469dbf10d64bc2ddac17e12487c2304cfa93b3262711702b8fad4";

/** Hash of the content, independent of how the checkout wrote the newlines. */
function normalisedSha256(text: string): string {
  return createHash("sha256").update(text.replace(/\r\n/g, "\n"), "utf8").digest("hex");
}

describe("holy-sheet.schema.json is in sync with the PHP twin", () => {
  const path = new URL("../src/holy-sheet.schema.json", import.meta.url);

  it("matches the shared checksum", () => {
    expect(normalisedSha256(readFileSync(path, "utf8"))).toBe(SHARED_SCHEMA_SHA256);
  });

  it("is valid JSON with the fields a tool definition needs", () => {
    // A checksum alone would happily pin a corrupt file. This is the assertion
    // that the thing being pinned is still the thing we mean.
    const schema = JSON.parse(readFileSync(path, "utf8")) as Record<string, unknown>;

    expect(schema.$schema).toBeTypeOf("string");
    expect(schema.type).toBe("object");
    expect(schema.properties).toBeTypeOf("object");
  });
});
