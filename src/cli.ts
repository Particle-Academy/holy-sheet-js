#!/usr/bin/env node
/** The `holy-sheet` bin. All behaviour lives in `cli-run.ts`, which the tests import. */
import { run } from "./cli-run";

run(process.argv.slice(2), {
  stdout: (text) => process.stdout.write(text),
  stderr: (text) => process.stderr.write(text),
}).then((code) => {
  process.exitCode = code;
});
