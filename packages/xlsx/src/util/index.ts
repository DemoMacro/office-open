/**
 * Convert a 1-based column number to Excel column letter(s).
 * 1 → "A", 26 → "Z", 27 → "AA", 28 → "AB"
 *
 * Memoized: sheet building calls this per cell (millions of times over a few
 * dozen distinct columns) — the cache turns each hit into an array read.
 */
const COLUMN_LETTERS: string[] = [];

export function columnToLetter(col: number): string {
  const cached = COLUMN_LETTERS[col];
  if (cached !== undefined) return cached;
  let result = "";
  let n = col;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  COLUMN_LETTERS[col] = result;
  return result;
}

/**
 * Convert Excel column letter(s) to a 1-based column number.
 * "A" → 1, "Z" → 26, "AA" → 27
 */
export function letterToColumn(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return col;
}

const A1_CELL = /^([A-Z]+)(\d+)$/;

/**
 * Parse a single A1 cell reference ("B12") into 1-based column/row numbers.
 * Returns undefined when the reference is not letters-then-digits.
 */
export function parseA1Cell(ref: string): { col: number; row: number } | undefined {
  const m = ref.match(A1_CELL);
  return m ? { col: letterToColumn(m[1]!), row: parseInt(m[2]!, 10) } : undefined;
}

/**
 * Convert a JavaScript Date to an Excel serial number.
 * Excel epoch: January 1, 1900 = 1 (with the 1900 leap year bug).
 */
export function dateToSerialNumber(date: Date): number {
  // Excel treats 1900 as a leap year (bug inherited from Lotus 1-2-3).
  // The epoch is effectively December 30, 1899 = 0, fixed in UTC so the
  // serial does not shift with the host timezone.
  const epochMs = Date.UTC(1899, 11, 30);
  return (date.getTime() - epochMs) / 86400000;
}

/**
 * Legacy 16-bit XOR password hash used by xlsx sheet/workbook protection
 * (the pre-Agile encryption hash). Shared by worksheet and workbook parts.
 */
export function hashPassword(password: string): string {
  let hash = 0;
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i);
    hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
    hash ^= c;
    hash = hash & 0x4000 ? hash ^ 0x1 : hash;
  }
  hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
  hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
  hash ^= password.length;
  return hash.toString(16).toUpperCase().padStart(4, "0");
}
