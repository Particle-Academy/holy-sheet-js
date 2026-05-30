/**
 * ISO string / Date → Excel serial number. Mirrors PHP `Writer\Format\DateConverter`.
 * Excel epoch is anchored to 1899-12-30 (accommodating the legacy 1900 leap-year bug).
 */

function parseToUtcMs(value: string): number | null {
  const s = value.trim();
  if (s === "") return null;
  const m =
    /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2})(?:\.\d+)?)?)?\s*(Z|[+-]\d{2}:?\d{2})?$/.exec(
      s,
    );
  if (m) {
    const [, y, mo, d, h, mi, se, tz] = m;
    let ms = Date.UTC(+y!, +mo! - 1, +d!, h ? +h : 0, mi ? +mi : 0, se ? +se : 0);
    if (tz && tz !== "Z") {
      const sign = tz[0] === "-" ? -1 : 1;
      const clean = tz.slice(1).replace(":", "");
      const offMin = sign * (parseInt(clean.slice(0, 2), 10) * 60 + parseInt(clean.slice(2, 4) || "0", 10));
      ms -= offMin * 60000;
    }
    return ms;
  }
  const fallback = Date.parse(s);
  return Number.isNaN(fallback) ? null : fallback;
}

export const DateConverter = {
  toSerial(value: string | Date, includeTime = false): number {
    const ts = value instanceof Date ? value.getTime() : parseToUtcMs(value);
    if (ts === null) return 0;
    const epoch = Date.UTC(1899, 11, 30, 0, 0, 0);
    let days = (ts - epoch) / 86400000;
    if (!includeTime) days = Math.floor(days);
    return days;
  },
};
