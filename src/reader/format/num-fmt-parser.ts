import type { DisplayFormat } from "../../schema/types";

export interface NumFmtParsed {
  displayFormat: DisplayFormat;
  decimals?: number;
  currency?: string;
}

/** Reverse of NumFmtBuilder. Mirrors PHP `Reader\Format\NumFmtParser`. */
const BUILTIN: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

const SYMBOL_TO_ISO: Record<string, string> = {
  $: "USD",
  "€": "EUR",
  "£": "GBP",
  "¥": "JPY",
  "₹": "INR",
  "₩": "KRW",
};

export const NumFmtParser = {
  parse(formatCode: string | null): NumFmtParsed | null {
    if (formatCode === null) return null;
    const code = formatCode;

    if (code === "" || code === "General" || code === "@") {
      return { displayFormat: code === "@" ? "text" : "auto" };
    }

    if (looksLikeDate(code)) {
      return { displayFormat: hasTimeComponent(code) ? "datetime" : "date" };
    }

    if (code.includes("%")) {
      return { displayFormat: "percentage", decimals: decimalsAfter(code) };
    }

    const m = /^"([^"]+)"/u.exec(code);
    if (m) {
      const iso = SYMBOL_TO_ISO[m[1]!];
      const result: NumFmtParsed = { displayFormat: "currency", decimals: decimalsAfter(code) };
      if (iso !== undefined) result.currency = iso;
      return result;
    }

    if (/^#?,?#?#?0(\.0+)?$/.test(code.split(";")[0]!)) {
      return { displayFormat: "number", decimals: decimalsAfter(code) };
    }

    return null;
  },

  parseBuiltin(id: number): NumFmtParsed | null {
    const code = BUILTIN[id];
    if (code === undefined) return null;
    return NumFmtParser.parse(code);
  },
};

function looksLikeDate(code: string): boolean {
  const stripped = code.replace(/"[^"]*"|\[[^\]]*\]/gu, "");
  return /[ymd]/i.test(stripped);
}

function hasTimeComponent(code: string): boolean {
  const stripped = code.replace(/"[^"]*"|\[[^\]]*\]/gu, "");
  return /[hHsS]/.test(stripped) || stripped.toLowerCase().includes("mm:");
}

function decimalsAfter(code: string): number {
  const section = code.split(";")[0]!;
  const m = /0\.(0+)/.exec(section);
  return m ? m[1]!.length : 0;
}
