import { readUint16 } from "./binary";
import type { LegacyParagraphProperties } from "./models";
import { iterateProperties } from "./property-iterator";
const ALIGNMENT_VALUES = ["left", "center", "right", "both"] as const;

export function parseParagraphProperties(papxText: Uint8Array): LegacyParagraphProperties {
  const properties: LegacyParagraphProperties = {};
  iterateProperties(papxText, (opcode, operandOffset, length) => {
    const operation = opcode & 0x00ff;
    const signed = () =>
      (readUint16(papxText, operandOffset, "Invalid DOC paragraph properties") << 16) >> 16;
    if (operation === 0x03 && length === 1) {
      properties.alignment = ALIGNMENT_VALUES[papxText[operandOffset]! & 3];
    } else if (operation === 0x0f || operation === 0x5e) {
      properties.indentLeft = signed();
    } else if (operation === 0x0e || operation === 0x5d) {
      properties.indentRight = signed();
    } else if (operation === 0x11 || operation === 0x60) {
      properties.firstLine = Math.max(0, signed());
    } else if (operation === 0x13 && length === 2) {
      properties.spacingBefore = signed();
    } else if (operation === 0x14 && length === 2) {
      properties.spacingAfter = signed();
    } else if (operation === 0x12 && length >= 4) {
      properties.spacingLine = readUint16(
        papxText,
        operandOffset + 2,
        "Invalid DOC paragraph properties",
      );
    } else if (operation === 0x40 && length === 1) {
      properties.outlineLevel = papxText[operandOffset]! & 15;
    } else if (operation === 0x16 && length === 1) {
      properties.inTable = papxText[operandOffset] !== 0;
    } else if (operation === 0x17 && length === 1) {
      properties.rowEnd = papxText[operandOffset] !== 0;
    } else if (operation === 0x08 && length > 2) {
      parseTableProperties(
        papxText.subarray(operandOffset + 1, operandOffset + length),
        properties,
      );
    }
  });
  return properties;
}

function parseTableProperties(tap: Uint8Array, properties: LegacyParagraphProperties): void {
  if (tap.byteLength < 4) return;
  const cellCount = tap[0]!;
  if (cellCount === 0 || 2 + (cellCount + 1) * 2 > tap.byteLength) return;
  properties.columnWidths = [];
  let previous = tap[2]! | (tap[3]! << 8);
  for (let index = 0; index < cellCount; index += 1) {
    const offset = 2 + (index + 1) * 2;
    const boundary = tap[offset]! | (tap[offset + 1]! << 8);
    properties.columnWidths.push(Math.max(0, boundary - previous));
    previous = boundary;
  }
  properties.inTable = true;
  const tcOffset = 2 + (cellCount + 1) * 2;
  properties.cellMerges = [];
  for (let index = 0; tcOffset + index + 1 < tap.byteLength && index < cellCount; index += 1) {
    const flags = tap[tcOffset + index]! | (tap[tcOffset + index + 1]! << 8);
    properties.cellMerges.push({
      restart: (flags & 0x0002) !== 0,
      continue: (flags & 0x0001) !== 0,
    });
  }
}
