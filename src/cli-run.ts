/**
 * `npx holy-sheet` (the `bin`). One command today, mirroring PHP `bin/holy-sheet`:
 *
 *     holy-sheet template:convert <input> <output> [--to=document|template] [--force]
 *
 * Exit codes: 0 converted, 1 refused (the reason is printed), 2 usage.
 *
 * Node-only, and never imported by the isomorphic index.
 *
 * @internal
 */
import { existsSync, realpathSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { randomBytes } from "node:crypto";
import { Agent } from "./agent";
import { ArchiveRefused, OfficeTemplateRefused, UnsupportedFormatException } from "./exceptions";
import { readFileBounded } from "./reader/file-input";
import { SourceArchive } from "./zip/source-archive";
import { sourceKind, type SourceKind } from "./zip/office-template";

export interface CliIo {
  stdout: (text: string) => void;
  stderr: (text: string) => void;
}

const USAGE = "Usage: holy-sheet template:convert <input> <output> [--to=document|template] [--force]\n"
  + "  Converts an .xltx/.ots template to its .xlsx/.ods twin, or back.\n"
  + "  Without --to it flips whichever the input is.\n";

export async function run(args: string[], io: CliIo): Promise<number> {
  const [command, ...rest] = args;
  const positional: string[] = [];
  let to: string | null = null;
  let force = false;
  for (const arg of rest) {
    if (arg.startsWith("--to=")) to = arg.slice(5);
    else if (arg === "--force") force = true;
    else if (arg.startsWith("--")) {
      io.stderr(`Unknown option ${arg}\n${USAGE}`);
      return 2;
    } else positional.push(arg);
  }
  if (command !== "template:convert" || positional.length !== 2) {
    io.stderr(USAGE);
    return 2;
  }
  const [input, output] = positional as [string, string];

  if (existsSync(output) && !force) {
    io.stderr(`Refused: ${output} already exists (pass --force to replace it).\n`);
    return 1;
  }
  if (existsSync(input) && existsSync(output) && realpathSync(input) === realpathSync(output)) {
    io.stderr("Refused: the output is the input.\n");
    return 1;
  }
  if (!existsSync(input) || !statSync(input).isFile()) {
    io.stderr(`Refused: ${input} is not a file.\n`);
    return 1;
  }

  let converted: Uint8Array;
  try {
    converted = Agent.convertTemplate(await readFileBounded(input), to as SourceKind | null);
  } catch (error) {
    if (error instanceof OfficeTemplateRefused || error instanceof ArchiveRefused) {
      io.stderr(`Refused (${error.reason}): ${error.message}\n`);
      return 1;
    }
    if (error instanceof UnsupportedFormatException) {
      io.stderr(`Refused (unsupported): ${error.message}\n`);
      return 1;
    }
    throw error;
  }

  // Written beside the target and renamed, so a failure leaves no half file.
  const temp = `${output}.tmp-${randomBytes(4).toString("hex")}`;
  try {
    writeFileSync(temp, converted);
    renameSync(temp, output);
  } catch {
    try { unlinkSync(temp); } catch { /* already gone */ }
    io.stderr(`Could not write ${output}.\n`);
    return 1;
  }
  io.stdout(`Wrote ${output} (${sourceKind(SourceArchive.open(converted))}).\n`);
  return 0;
}
