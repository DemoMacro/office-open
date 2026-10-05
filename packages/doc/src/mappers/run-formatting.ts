import type { RunOptions } from "@office-open/docx";

import { readUint16 } from "./binary";
import { iterateProperties } from "./property-iterator";

const COLOR_INDEXES = [
  "auto",
  "000000",
  "0000ff",
  "00ffff",
  "00ff00",
  "ff00ff",
  "ff0000",
  "ffff00",
  "ffffff",
  "000080",
  "008080",
  "008000",
  "800080",
  "800000",
  "808000",
  "808080",
  "c0c0c0",
] as const;
export function parseCharacterProperties(chpxText: Uint8Array): RunOptions {
  const properties: RunOptions = {};
  iterateProperties(chpxText, (opcode, operandOffset, length) => {
    if (opcode === 0x0835 && length === 1)
      properties.bold = chpxText[operandOffset]! === 1 || chpxText[operandOffset]! === 128;
    if (opcode === 0x0836 && length === 1)
      properties.italic = chpxText[operandOffset]! === 1 || chpxText[operandOffset]! === 128;
    if (opcode === 0x0837 && length === 1)
      properties.strike = chpxText[operandOffset]! === 1 || chpxText[operandOffset]! === 128;
    if (opcode === 0x4a43 && length === 2) {
      properties.size = readUint16(chpxText, operandOffset, "Invalid DOC character properties") / 2;
    }
    if (opcode === 0x6a03 && length === 2) {
      const colorIndex = chpxText[operandOffset]!;
      if (colorIndex < COLOR_INDEXES.length) properties.color = COLOR_INDEXES[colorIndex]!;
    }
    if (opcode === 0x2a47 && length === 1) {
      const style = chpxText[operandOffset]!;
      properties.underline =
        style === 0 || style === undefined
          ? undefined
          : { type: style === 2 ? "words" : style === 3 ? "double" : "single" };
    }
    if (opcode === 0x2a48 && length === 1) {
      properties.highlight = chpxText[operandOffset] === 0 ? undefined : "yellow";
    }
    if (opcode === 0x4a4f && length === 2) {
      properties.font = `font-${readUint16(chpxText, operandOffset, "Invalid DOC character properties")}`;
    }
  });
  return properties;
}
