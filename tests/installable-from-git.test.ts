import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";

/**
 * `dist/` is gitignored, so a git install (`npm install
 * github:Particle-Academy/holy-sheet-js#<sha>`) only has an entry point if npm
 * builds one. npm runs `prepare` (and only `prepare`) for a git dependency,
 * after installing its devDependencies. Without it the package installs with
 * no `dist/index.js`, and every import fails with ERR_MODULE_NOT_FOUND. The
 * lab installs release candidates exactly this way, which is how this was
 * found (2026-11-19).
 */
describe("installable from a git ref", () => {
    const pkg = JSON.parse(readFileSync(join(__dirname, "..", "package.json"), "utf8"));

    it("builds dist on prepare, the one script npm runs for a git dependency", () => {
        expect(pkg.scripts?.prepare).toBe(pkg.scripts?.build);
        expect(pkg.scripts?.build).toBe("tsup");
    });

    it("ships the dist the build writes", () => {
        expect(pkg.files).toContain("dist");
        expect(pkg.main ?? pkg.exports).toBeTruthy();
    });
});
