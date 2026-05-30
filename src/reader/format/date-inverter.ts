/** Excel serial → ISO-8601 string. Mirrors PHP `Reader\Format\DateInverter`. */
export const DateInverter = {
  toIso(serial: number, includeTime = false): string {
    const epoch = Date.UTC(1899, 11, 30, 0, 0, 0);
    const ms = epoch + Math.round(serial * 86400) * 1000;
    const d = new Date(ms);
    const pad = (n: number) => String(n).padStart(2, "0");
    const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    if (!includeTime) return date;
    return `${date}T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`;
  },
};
