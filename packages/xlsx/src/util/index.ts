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
