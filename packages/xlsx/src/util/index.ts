/**
 * Convert a 1-based column number to Excel column letter(s).
 * 1 → "A", 26 → "Z", 27 → "AA", 28 → "AB"
 */
export function columnToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

/**
 * Convert a JavaScript Date to an Excel serial number.
 * Excel epoch: January 1, 1900 = 1 (with the 1900 leap year bug).
 */
export function dateToSerialNumber(date: Date): number {
  // Excel treats 1900 as a leap year (bug inherited from Lotus 1-2-3).
  // The epoch is effectively December 30, 1899 = 0.
  const epoch = new Date(1899, 11, 30);
  const msPerDay = 86400000;
  const diff = date.getTime() - epoch.getTime();
  return diff / msPerDay;
}
