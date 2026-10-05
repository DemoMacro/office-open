import { DocParseError } from "../errors";
import { readUint16, requireRange } from "./binary";
function operandLength(opcode: number, bytes: Uint8Array, offset: number): number {
  switch ((opcode >> 13) & 0b111) {
    case 0:
    case 1:
      return 1;
    case 2:
    case 4:
    case 5:
      return 2;
    case 3:
      return 4;
    case 7:
      return 3;
    case 6:
      requireRange(bytes, offset, 1, "Invalid DOC properties: truncated variable operand");
      return 1 + bytes[offset]!;
    default:
      throw new DocParseError("Invalid DOC properties: unknown operand form");
  }
}

export function iterateProperties(
  grpprl: Uint8Array,
  visit?: (opcode: number, operandOffset: number, operandLength: number) => void,
): void {
  let offset = 0;
  while (offset < grpprl.byteLength) {
    if (grpprl.byteLength - offset < 2) {
      if (grpprl.byteLength - offset === 1) break;
      throw new DocParseError("Invalid DOC properties: truncated property opcode");
    }
    const opcode = readUint16(grpprl, offset, "Invalid DOC properties: truncated property opcode");
    const length = operandLength(opcode, grpprl, offset + 2);
    const operandOffset = offset + 2;
    requireRange(
      grpprl,
      operandOffset,
      length,
      "Invalid DOC properties: truncated property operand",
    );
    visit?.(opcode, operandOffset, length);
    offset = operandOffset + length;
  }
}
