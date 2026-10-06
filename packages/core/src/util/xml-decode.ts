/**
 * Decode an XML byte payload according to its BOM or declared encoding.
 *
 * OOXML packages in the wild may contain UTF-16 parts even though newly
 * generated parts are UTF-8.
 *
 * @module
 */

export function decodeXmlBytes(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe)
    return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff)
    return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] !== 0 && bytes[1] === 0)
    return new TextDecoder("utf-16le").decode(bytes);
  if (bytes.length >= 2 && bytes[0] === 0 && bytes[1] !== 0)
    return new TextDecoder("utf-16be").decode(bytes);

  const declaration = new TextDecoder("latin1").decode(
    bytes.subarray(0, Math.min(120, bytes.length)),
  );
  if (/encoding\s*=\s*["'](?:utf-?16|unicode(?:feff)?)["']/i.test(declaration)) {
    if (bytes[0] === 0 && bytes[1] !== 0) return new TextDecoder("utf-16be").decode(bytes);
    return new TextDecoder("utf-16le").decode(bytes);
  }
  return new TextDecoder().decode(bytes);
}
