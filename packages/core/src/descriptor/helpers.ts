/**
 * OOXML-specific encode/decode helpers for descriptors.
 *
 * XML traversal helpers (findChild, etc.) are in @office-open/xml utils.
 * This file only contains OOXML value encoding/decoding.
 *
 * @module
 */

// ── Boolean encode/decode (CT_OnOff pattern) ──

/** Encode boolean for CT_OnOff: true → omit val, false → "0". */
export const boolEncode = (v: boolean | undefined): string | undefined => {
  if (v === undefined) return undefined;
  return v ? undefined : "0";
};

/** Decode CT_OnOff: absent or "true"/"1" → true, "0"/"false" → false. */
export const boolDecode = (raw: string): boolean => raw !== "0" && raw !== "false";

// ── Enum mapping ──

/** Create an enum encoder from a JS↔XML mapping. */
export const enumEncode =
  (map: Record<string, string>) =>
  (v: string | undefined): string | undefined => {
    if (v === undefined) return undefined;
    return map[v] ?? v;
  };

/** Create an enum decoder from a JS↔XML mapping (inverted). */
export const enumDecode = (map: Record<string, string>) => {
  const inv = invertRecord(map);
  return (raw: string): string => inv[raw] ?? raw;
};

function invertRecord(map: Record<string, string>): Record<string, string> {
  const result: Record<string, string> = {};
  for (const key of Object.keys(map)) {
    result[map[key]] = key;
  }
  return result;
}
