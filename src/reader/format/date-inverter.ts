/**
 * Excel serial → ISO-8601 string. Mirrors PHP `Reader\Format\DateInverter`.
 *
 * A workbook that declares `<workbookPr date1904="1"/>` counts from 1904-01-01
 * instead (the old Mac default), so the same serial is 1462 days later;
 * `date1904` selects that epoch.
 */
export const DateInverter = {
  toIso(serial: number, includeTime = false, date1904 = false): string {
    const epoch = date1904 ? Date.UTC(1904, 0, 1, 0, 0, 0) : Date.UTC(1899, 11, 30, 0, 0, 0);
    const ms = epoch + Math.round(serial * 86400) * 1000;
    const d = new Date(ms);
    const pad = (n: number) => String(n).padStart(2, "0");
    const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    if (!includeTime) return date;
    return `${date}T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`;
  },
};
