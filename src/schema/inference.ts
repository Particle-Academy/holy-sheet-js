import { isNumericString } from "../util";
import type { BuilderOptions, ColumnSchema } from "./types";

/** Type inference for tabular data → column types. Mirrors PHP `Schema\Inference`. */
const SAMPLE_SIZE = 50;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const ISO_DATETIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$/;
const HEADER_INTEGER = /(^|[\s_])(count|qty|quantity|num|number|id|n)([\s_]|$)/i;
const HEADER_CURRENCY = /(price|amount|cost|revenue|fee|total|salary|budget|balance)/i;
const HEADER_PERCENT = /(rate|percent|growth|yoy|margin|share|ratio)/i;

type Any = any;

export const Inference = {
  detect(columnValues: unknown[], headerName: string, options: BuilderOptions = {}): ColumnSchema {
    const sample = nonNullSample(columnValues);

    if (sample.length === 0) return { header: headerName, type: "auto" };

    if (allBoolean(sample)) return { header: headerName, type: "boolean" };

    const allNum = allNumeric(sample);
    const allInRange = allNum && allInRange01(sample);
    const allInt = allNum && allInteger(sample);
    const allDate = allMatch(sample, ISO_DATE);
    const allDateTime = allMatch(sample, ISO_DATETIME);

    if (allDate) return { header: headerName, type: "date" };
    if (allDateTime) return { header: headerName, type: "datetime" };

    if (allNum) {
      if (HEADER_PERCENT.test(headerName) && allInRange) {
        return { header: headerName, type: "percent", decimals: detectDecimals(sample) || 1 };
      }
      if (HEADER_CURRENCY.test(headerName)) {
        return {
          header: headerName,
          type: "currency",
          currency: options.currency ?? "USD",
          decimals: detectDecimals(sample) ?? undefined,
        } as ColumnSchema;
      }
      if (HEADER_INTEGER.test(headerName) && allInt) {
        return { header: headerName, type: "integer" };
      }
      if (allInt) return { header: headerName, type: "integer" };

      const col: ColumnSchema = { header: headerName, type: "number" };
      const decimals = detectDecimals(sample);
      if (decimals !== null && decimals > 0) col.decimals = decimals;
      return col;
    }

    if (allStringish(sample)) return { header: headerName, type: "string" };

    return { header: headerName, type: "auto" };
  },
};

function nonNullSample(values: unknown[]): unknown[] {
  const out: unknown[] = [];
  for (const v of values) {
    if (v === null || v === undefined) continue;
    out.push(v);
    if (out.length >= SAMPLE_SIZE) break;
  }
  return out;
}

function allBoolean(sample: unknown[]): boolean {
  return sample.every((v) => typeof v === "boolean");
}

function allNumeric(sample: unknown[]): boolean {
  return sample.every((v) => {
    if (typeof v === "boolean") return false;
    if (typeof v === "number") return Number.isFinite(v);
    return typeof v === "string" && isNumericString(v);
  });
}

function allInteger(sample: unknown[]): boolean {
  return sample.every((v) => {
    if (typeof v === "number") return Number.isInteger(v);
    return typeof v === "string" && /^-?\d+$/.test(v);
  });
}

function allInRange01(sample: unknown[]): boolean {
  return sample.every((v) => {
    const f = Number(v as Any);
    return f >= 0 && f <= 1;
  });
}

function detectDecimals(sample: unknown[]): number | null {
  let max = 0;
  let sawFloat = false;
  for (const v of sample) {
    const s = typeof v === "string" ? v : String(v);
    if (s.includes(".")) {
      sawFloat = true;
      const frac = (s.split(".", 2)[1] ?? "").replace(/0+$/, "");
      if (frac.length > max) max = frac.length;
    }
  }
  if (!sawFloat) return 0;
  return Math.max(max, 2);
}

function allMatch(sample: unknown[], regex: RegExp): boolean {
  return sample.every((v) => typeof v === "string" && regex.test(v));
}

function allStringish(sample: unknown[]): boolean {
  return sample.every((v) => typeof v === "string");
}
