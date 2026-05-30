/**
 * XML escaping, matching the PHP writer's `escape()`:
 * strips disallowed control characters, then escapes the 5 XML metacharacters
 * (`htmlspecialchars(..., ENT_XML1 | ENT_QUOTES)`).
 */

// PHP: /[\x00-\x08\x0B\x0C\x0E-\x1F]/u  — control chars not permitted in XML 1.0.
const CONTROL_CHARS = new RegExp("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]", "g");

export function xmlEscape(s: string): string {
  return s
    .replace(CONTROL_CHARS, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
