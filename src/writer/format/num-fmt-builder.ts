import type { CellFormat } from "../../workbook/cell-format";

/**
 * Builds the Excel number-format code for a CellFormat. Returns null when no
 * numFmt is needed. Mirrors PHP `Writer\Format\NumFmtBuilder`.
 */
export const NumFmtBuilder = {
  build(format: CellFormat): string | null {
    const df = format.displayFormat;
    const decimals = format.decimals;

    if (df === null || df === "auto" || df === "text") return null;

    switch (df) {
      case "number":
        return numberFormat(decimals);
      case "percentage":
        return percentFormat(decimals);
      case "currency":
        return currencyFormat(format.currency, decimals);
      case "date":
        return "yyyy-mm-dd";
      case "datetime":
        return "yyyy-mm-dd hh:mm:ss";
      default:
        return null;
    }
  },
};

function numberFormat(decimals: number | null): string {
  if (decimals === null || decimals <= 0) return "#,##0";
  return "#,##0." + "0".repeat(decimals);
}

function percentFormat(decimals: number | null): string {
  const d = decimals ?? 1;
  if (d <= 0) return "0%";
  return "0." + "0".repeat(d) + "%";
}

function currencyFormat(currency: string | null, decimals: number | null): string {
  const symbol = currencySymbol(currency ?? "USD");
  const d = decimals ?? 2;
  const body = d <= 0 ? "#,##0" : "#,##0." + "0".repeat(d);
  return `"${symbol}"${body};-"${symbol}"${body}`;
}

function currencySymbol(iso: string): string {
  switch (iso.toUpperCase()) {
    case "USD":
      return "$";
    case "EUR":
      return "€";
    case "GBP":
      return "£";
    case "JPY":
      return "¥";
    case "CNY":
      return "¥";
    case "INR":
      return "₹";
    case "AUD":
      return "A$";
    case "CAD":
      return "C$";
    case "CHF":
      return "CHF";
    case "KRW":
      return "₩";
    default:
      return iso + " ";
  }
}
