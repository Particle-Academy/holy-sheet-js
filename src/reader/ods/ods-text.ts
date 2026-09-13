import type { XmlNode } from "../xml";
import { NS } from "./ns";

/**
 * The text of a cell or annotation: its paragraphs, one line each. Mirrors PHP
 * `Reader\Ods\OdsText`.
 *
 * A paragraph is mixed content — runs, links, `text:s` for extra spaces,
 * `text:tab`, `text:line-break` — and its meaning depends on ORDER, so it is
 * walked through `content`. Inside a paragraph, elements are matched by local
 * name, as the PHP reader matches them.
 *
 * White space follows ODF 1.3 §6.1.2, the way LibreOffice applies it: a run of
 * space, tab, CR and LF characters is ONE space, and white space at the start of
 * a paragraph or straight after another collapsed space is dropped. `text:s` is
 * never collapsed.
 */

/** Elements whose text is not part of the paragraph (a note, a comment). */
const SKIPPED = new Set(["annotation", "note"]);

interface State {
  out: string;
  ignoreSpace: boolean;
}

/** The `text:p` / `text:h` children joined with "\n", or null when there are none. */
export function paragraphs(node: XmlNode): string | null {
  const lines: string[] = [];
  for (const c of node.children) {
    if (c.ns === NS.TEXT && (c.name === "p" || c.name === "h")) lines.push(paragraph(c));
  }
  return lines.length === 0 ? null : lines.join("\n");
}

export function paragraph(p: XmlNode): string {
  const state: State = { out: "", ignoreSpace: true };
  walk(p, state);
  return state.out;
}

function walk(node: XmlNode, state: State): void {
  for (const item of node.content ?? []) {
    if (typeof item === "string") {
      appendCharacters(state, item);
      continue;
    }
    if (SKIPPED.has(item.name)) continue;

    if (item.name === "s") {
      const c = item.attrs["c"];
      const count = c !== undefined && /^\d+$/.test(c) ? parseInt(c, 10) : 1;
      state.out += " ".repeat(Math.max(1, count));
      state.ignoreSpace = false;
    } else if (item.name === "tab") {
      state.out += "\t";
      state.ignoreSpace = false;
    } else if (item.name === "line-break") {
      state.out += "\n";
      state.ignoreSpace = false;
    }
    walk(item, state);
  }
}

function appendCharacters(state: State, chars: string): void {
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i]!;
    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      if (!state.ignoreSpace) {
        state.out += " ";
        state.ignoreSpace = true;
      }
      continue;
    }
    state.out += c;
    state.ignoreSpace = false;
  }
}
