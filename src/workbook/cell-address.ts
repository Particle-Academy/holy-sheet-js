/**
 * A1 cell address utilities — column-letter ↔ 0-based index conversion,
 * address parsing. Mirrors PHP `Workbook\CellAddress`.
 */
export const CellAddress = {
  /** 0-based column index → Excel-style letters. */
  letter(index: number): string {
    if (index < 0) {
      throw new Error(`[holy-sheet] column index must be ≥ 0, got ${index}`);
    }
    let letters = "";
    let n = index;
    do {
      letters = String.fromCharCode(65 + (n % 26)) + letters;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return letters;
  },

  /** Column-letters → 0-based index. */
  index(letters: string): number {
    const up = letters.trim().toUpperCase();
    if (up === "" || !/^[A-Z]+$/.test(up)) {
      throw new Error(`[holy-sheet] invalid column letters: '${up}'`);
    }
    let idx = 0;
    for (let i = 0; i < up.length; i++) {
      idx = idx * 26 + (up.charCodeAt(i) - 64);
    }
    return idx - 1;
  },

  /** Parse an A1 address into [columnIndex, rowNumber] (1-based row), or null. */
  parse(address: string): [number, number] | null {
    const m = /^([A-Z]+)(\d+)$/.exec(address.trim().toUpperCase());
    if (!m) return null;
    return [CellAddress.index(m[1]!), parseInt(m[2]!, 10)];
  },
};
