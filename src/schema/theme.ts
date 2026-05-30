import { CellFormat } from "../workbook/cell-format";

/** Theme presets — pre-baked CellFormat sets. Mirrors PHP `Schema\Theme`. */
export class Theme {
  constructor(readonly key: string) {}

  headerFormat(): CellFormat | null {
    switch (this.key) {
      case "default":
      case "business":
        return new CellFormat({
          bold: true,
          color: "#FFFFFF",
          backgroundColor: this.key === "business" ? "#1F2937" : "#374151",
        });
      case "minimal":
        return new CellFormat({ bold: true, borderBottom: "#000000" });
      default:
        return null;
    }
  }

  dataFormat(rowIndexZeroBased: number): CellFormat | null {
    if ((this.key === "default" || this.key === "business") && rowIndexZeroBased % 2 === 1) {
      return new CellFormat({ backgroundColor: "#F3F4F6" });
    }
    return null;
  }

  totalsFormat(): CellFormat | null {
    switch (this.key) {
      case "default":
      case "business":
      case "minimal":
        return new CellFormat({ bold: true, borderTop: "#000000" });
      default:
        return null;
    }
  }
}
