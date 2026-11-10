import { describe, expect, it } from "vitest";
import { Agent } from "../src";

/**
 * The agent guide names every public Agent method, or the method is on the
 * allowlist below with the reason it is left out. A method the guide never
 * names is a method no model reading the guide will use: 2.10 to 2.14 added
 * `rows`, `read`'s `sheet` and `limits`, border styles and data validation, and
 * the guide (last touched at 2.9.0) said nothing about any of them.
 *
 * "Names" means written as code: `Agent::x`, `Agent.x` or a backticked `x`, so
 * the English word "read" in a sentence does not count as documenting read().
 * The PHP and Python packages carry the same test over their own surfaces.
 */
const ALLOWLIST: Record<string, string> = {
  guide: "the guide itself: a model reading it already has it",
  version: "package metadata; it plays no part in building or reading a workbook",
};

const surface = (): string[] =>
  Object.entries(Agent)
    .filter(([, v]) => typeof v === "function")
    .map(([k]) => k);

const escape = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const names = (guide: string, method: string): boolean =>
  new RegExp("(Agent::|Agent\\.|`)" + escape(method) + "\\b").test(guide);

describe("agent guide covers the API", () => {
  it("names every public Agent method, or allowlists it with a reason", () => {
    const guide = Agent.guide();
    const missing = surface().filter((m) => !(m in ALLOWLIST) && !names(guide, m));
    expect(missing, `agent-guide/professional-workbooks.md never names: ${missing.join(", ")}`).toEqual([]);
  });

  it("allowlists only methods that exist, each with a reason, none the guide already names", () => {
    const guide = Agent.guide();
    const methods = surface();
    for (const [method, reason] of Object.entries(ALLOWLIST)) {
      expect(methods).toContain(method);
      expect(reason.trim()).not.toBe("");
      expect(names(guide, method), `${method} is named in the guide; drop it from the allowlist`).toBe(false);
    }
  });
});
