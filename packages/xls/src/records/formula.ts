import { EXCEL_ERROR_LITERALS, columnLabel } from "../shared/cell-reference";

export interface FormulaContext {
  baseRow: number;
  baseColumn: number;
  formulaVersion?: number;
  sheetNameByExternIndex?: (index: number) => string | undefined;
  definedNameByIndex?: (index: number) => string | undefined;
  arrayData?: Uint8Array;
}

const BINARY_OPERATORS: Record<number, string> = {
  0x03: "+",
  0x04: "-",
  0x05: "*",
  0x06: "/",
  0x07: "^",
  0x08: "&",
  0x09: "<",
  0x0a: "<=",
  0x0b: "=",
  0x0c: ">=",
  0x0d: ">",
  0x0e: "<>",
  0x0f: " ",
  0x10: ",",
  0x11: ":",
};

const FUNCTION_NAMES: Record<number, string> = {
  1: "IF",
  2: "ISNA",
  3: "ISERROR",
  4: "SUM",
  5: "AVERAGE",
  6: "MIN",
  7: "MAX",
  8: "COUNT",
  34: "ROUND",
  85: "CONCATENATE",
};

class FormulaReader {
  private readonly view: DataView;
  private cursor = 0;

  constructor(private readonly data: Uint8Array) {
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  }

  get position(): number {
    return this.cursor;
  }

  require(length: number): void {
    if (this.cursor + length > this.data.byteLength) {
      throw new XlsParseError("Invalid legacy XLS file: truncated formula");
    }
  }

  byte(): number {
    this.require(1);
    return this.data[this.cursor++]!;
  }

  uint16(): number {
    this.require(2);
    const value = this.view.getUint16(this.cursor, true);
    this.cursor += 2;
    return value;
  }

  uint32(): number {
    this.require(4);
    const value = this.view.getUint32(this.cursor, true);
    this.cursor += 4;
    return value;
  }

  float64(): number {
    this.require(8);
    const value = this.view.getFloat64(this.cursor, true);
    this.cursor += 8;
    return value;
  }

  bytes(length: number): Uint8Array {
    this.require(length);
    const value = this.data.subarray(this.cursor, this.cursor + length);
    this.cursor += length;
    return value;
  }
}

function absoluteCell(row: number, column: number, flags: number): string {
  const columnPart = `${(flags & 0x4000) === 0 ? "$" : ""}${columnLabel(column & 0x3fff)}`;
  const rowPart = `${(flags & 0x8000) === 0 ? "$" : ""}${row + 1}`;
  return `${columnPart}${rowPart}`;
}

function relativeCell(row: number, column: number, flags: number, context: FormulaContext): string {
  const rowDelta = viewSigned(row);
  const columnField = column & 0x3fff;
  const columnDelta = columnField >= 0x2000 ? columnField - 0x4000 : columnField;
  return absoluteCell(
    Math.max(0, context.baseRow + rowDelta),
    Math.max(0, context.baseColumn + columnDelta),
    flags,
  );
}

function viewSigned(value: number): number {
  return value >= 0x8000 ? value - 0x10000 : value;
}

function readString(reader: FormulaReader): string {
  const characterCount = reader.uint16();
  const flags = reader.byte();
  const highByte = (flags & 0x01) !== 0;
  const bytes = reader.bytes(characterCount * (highByte ? 2 : 1));
  return new TextDecoder(highByte ? "utf-16le" : "windows-1252").decode(bytes);
}

function biff5Address(reader: FormulaReader, context: FormulaContext): string {
  const rowField = reader.uint16();
  const column = reader.byte();
  const flags = ((rowField & 0x4000) === 0 ? 0x4000 : 0) | ((rowField & 0x8000) === 0 ? 0x8000 : 0);
  const row = rowField & 0x3fff;
  if ((rowField & 0xc000) === 0) return absoluteCell(row, column, flags);
  return relativeCell(row, column, flags, context);
}

function functionName(index: number): string {
  return FUNCTION_NAMES[index] ?? `FN_${index}`;
}

function readConstantArray(reader: FormulaReader): string {
  const columns = reader.byte() + 1;
  const rows = reader.byte() + 1;
  const values: string[][] = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row: string[] = [];
    for (let columnIndex = 0; columnIndex < columns; columnIndex += 1) {
      const type = reader.byte();
      if (type === 0x00) {
        reader.bytes(8);
        row.push("");
      } else if (type === 0x01) {
        row.push(String(reader.float64()));
      } else if (type === 0x02) {
        row.push(readString(reader));
      } else if (type === 0x04) {
        const value = reader.byte();
        reader.bytes(7);
        row.push(value === 0 ? "FALSE" : "TRUE");
      } else if (type === 0x10) {
        const error = reader.byte();
        reader.bytes(7);
        row.push(EXCEL_ERROR_LITERALS[error] ?? `#ERROR_${error}`);
      } else {
        throw new XlsParseError(
          `Invalid legacy XLS file: unsupported formula array constant 0x${type
            .toString(16)
            .padStart(2, "0")}`,
        );
      }
    }
    values.push(row);
  }
  return `{${values.map((row) => row.join(",")).join(";")}}`;
}

function decodeTokens(
  data: Uint8Array,
  context: FormulaContext,
): { formula: string; consumed: number } {
  const reader = new FormulaReader(data);
  const arrayReader = new FormulaReader(context.arrayData ?? new Uint8Array());
  const operands: string[] = [];
  let pendingFunction: { name: string; arguments: number } | undefined;
  let sharedFormulaReference = false;

  while (reader.position < data.byteLength) {
    const token = reader.byte();
    switch (token) {
      case 0x03:
      case 0x04:
      case 0x05:
      case 0x06:
      case 0x07:
      case 0x08:
      case 0x09:
      case 0x0a:
      case 0x0b:
      case 0x0c:
      case 0x0d:
      case 0x0e:
      case 0x0f:
      case 0x10:
      case 0x11: {
        const right = operands.pop();
        const left = operands.pop();
        if (left === undefined || right === undefined)
          throw new XlsParseError("Invalid legacy XLS file: malformed formula operation");
        operands.push(`${left}${BINARY_OPERATORS[token]}${right}`);
        break;
      }
      case 0x13:
      case 0x12: {
        const operand = operands.pop();
        if (operand === undefined)
          throw new XlsParseError("Invalid legacy XLS file: malformed unary formula");
        operands.push(token === 0x13 ? `-${operand}` : `+${operand}`);
        break;
      }
      case 0x01: {
        reader.uint16();
        reader.uint16();
        sharedFormulaReference = true;
        break;
      }
      case 0x15: {
        const operand = operands.pop();
        if (operand === undefined)
          throw new XlsParseError("Invalid legacy XLS file: malformed parenthesized formula");
        operands.push(`(${operand})`);
        break;
      }
      case 0x1c: {
        const error = reader.byte();
        const message = EXCEL_ERROR_LITERALS[error];
        if (!message)
          throw new XlsParseError(
            `Invalid legacy XLS file: unsupported formula error 0x${error
              .toString(16)
              .padStart(2, "0")}`,
          );
        operands.push(message);
        break;
      }
      case 0x29:
      case 0x49:
      case 0x69: {
        const length = reader.uint16();
        const subexpression = decodeTokens(reader.bytes(length), context).formula;
        operands.push(subexpression);
        break;
      }
      case 0x20:
      case 0x40:
      case 0x60: {
        reader.bytes(7);
        if (!context.arrayData || context.arrayData.byteLength === 0) {
          operands.push('""');
          break;
        }
        operands.push(readConstantArray(arrayReader));
        break;
      }
      case 0x19: {
        const attribute = reader.byte();
        if (attribute === 0x04) {
          const row = reader.uint16();
          const column = reader.uint16();
          operands.push(reference(row, column & 0x3fff));
        } else {
          reader.uint16();
        }
        break;
      }
      case 0x21: {
        const index = reader.uint16();
        const argumentCount = reader.byte();
        pendingFunction = { name: functionName(index), arguments: argumentCount };
        break;
      }
      case 0x22: {
        const argumentCount = reader.byte();
        const index = reader.uint16();
        pendingFunction = { name: functionName(index), arguments: argumentCount };
        break;
      }
      case 0x23:
      case 0x43:
      case 0x63: {
        const index = reader.uint16() - 1;
        reader.uint16();
        operands.push(context.definedNameByIndex?.(index) ?? `#NAME_${index + 1}`);
        break;
      }
      case 0x24:
      case 0x44:
      case 0x64: {
        const row = reader.uint16();
        const column = reader.uint16();
        operands.push(absoluteCell(row, column, column >> 8));
        break;
      }
      case 0x25:
      case 0x45:
      case 0x65: {
        const firstRow = reader.uint16();
        const lastRow = reader.uint16();
        const firstColumn = reader.uint16();
        const lastColumn = reader.uint16();
        operands.push(
          `${absoluteCell(firstRow, firstColumn, firstColumn >> 8)}:${absoluteCell(
            lastRow,
            lastColumn,
            lastColumn >> 8,
          )}`,
        );
        break;
      }
      case 0x3a:
      case 0x5a:
      case 0x7a: {
        const externIndex = reader.uint16();
        if (context.formulaVersion === 5) {
          reader.bytes(8);
          reader.uint16();
          reader.uint16();
          const address = biff5Address(reader, context);
          const sheet = context.sheetNameByExternIndex?.(viewSigned(externIndex)) ?? "#REF";
          operands.push(sheet ? `${sheet}!${address}` : address);
        } else {
          const row = reader.uint16();
          const column = reader.uint16();
          const sheet = context.sheetNameByExternIndex?.(externIndex) ?? "#REF";
          operands.push(`${sheet}!${absoluteCell(row, column, column >> 8)}`);
        }
        break;
      }
      case 0x3b:
      case 0x5b:
      case 0x7b: {
        const externIndex = reader.uint16();
        if (context.formulaVersion === 5) {
          reader.bytes(8);
          reader.uint16();
          reader.uint16();
          const first = biff5Address(reader, context);
          const last = biff5Address(reader, context);
          const sheet = context.sheetNameByExternIndex?.(viewSigned(externIndex)) ?? "#REF";
          operands.push(sheet ? `${sheet}!${first}:${last}` : `${first}:${last}`);
        } else {
          const firstRow = reader.uint16();
          const lastRow = reader.uint16();
          const firstColumn = reader.uint16();
          const lastColumn = reader.uint16();
          const sheet = context.sheetNameByExternIndex?.(externIndex) ?? "#REF";
          operands.push(
            `${sheet}!${absoluteCell(firstRow, firstColumn, firstColumn >> 8)}:${absoluteCell(
              lastRow,
              lastColumn,
              lastColumn >> 8,
            )}`,
          );
        }
        break;
      }
      case 0x17: {
        const value = readString(reader);
        operands.push(`"${value.replaceAll('"', '""')}"`);
        break;
      }
      case 0x3c:
      case 0x5c:
      case 0x7c: {
        const externIndex = reader.uint16();
        if (context.formulaVersion === 5) {
          reader.bytes(8);
          reader.uint16();
          reader.uint16();
          reader.bytes(3);
        } else {
          reader.bytes(2);
        }
        const sheet = context.sheetNameByExternIndex?.(externIndex) ?? "#REF";
        operands.push(`${sheet}!#REF!`);
        break;
      }
      case 0x3d:
      case 0x5d:
      case 0x7d: {
        const externIndex = reader.uint16();
        if (context.formulaVersion === 5) {
          reader.bytes(8);
          reader.uint16();
          reader.uint16();
          reader.bytes(6);
        } else {
          reader.bytes(8);
        }
        const sheet = context.sheetNameByExternIndex?.(externIndex) ?? "#REF";
        operands.push(`${sheet}!#REF!`);
        break;
      }
      case 0x41:
      case 0x61: {
        const argumentCount = reader.byte();
        const index = reader.uint16();
        pendingFunction = { name: functionName(index), arguments: argumentCount };
        break;
      }
      case 0x1d: {
        operands.push(reader.byte() === 0 ? "FALSE()" : "TRUE()");
        break;
      }
      case 0x42:
      case 0x62: {
        const error = reader.byte();
        operands.push(`#VALUE_${error}`);
        break;
      }
      case 0x1e: {
        operands.push(String(reader.uint16()));
        break;
      }
      case 0x1f: {
        operands.push(String(reader.float64()));
        break;
      }
      case 0x4c:
      case 0x6c: {
        const row = reader.uint16();
        const column = reader.uint16();
        operands.push(relativeCell(row, column, column >> 8, context));
        break;
      }
      default:
        const operatorToken = token >= 0x20 ? token & 0x1f : token;
        if (BINARY_OPERATORS[operatorToken] !== undefined) {
          const right = operands.pop();
          const left = operands.pop();
          if (left === undefined || right === undefined) {
            throw new XlsParseError("Invalid legacy XLS file: malformed formula operation");
          }
          operands.push(`${left}${BINARY_OPERATORS[operatorToken]}${right}`);
          break;
        }
        throw new XlsParseError(
          `Invalid legacy XLS file: unsupported formula token 0x${token.toString(16).padStart(2, "0")}`,
        );
    }

    if (pendingFunction && operands.length >= pendingFunction.arguments) {
      const argumentsText = operands
        .splice(operands.length - pendingFunction.arguments, pendingFunction.arguments)
        .join(",");
      operands.push(`${pendingFunction.name}(${argumentsText})`);
      pendingFunction = undefined;
    }
  }

  if (
    (operands.length !== 1 && !(sharedFormulaReference && operands.length === 0)) ||
    pendingFunction
  ) {
    throw new XlsParseError("Invalid legacy XLS file: malformed formula token stream");
  }
  return { formula: operands[0] ?? "", consumed: data.byteLength };
}

export function decodeFormula(
  data: Uint8Array,
  context: FormulaContext,
): { formula: string; consumed: number } {
  const result = decodeTokens(data, context);
  return { ...result, consumed: result.consumed + (context.arrayData?.byteLength ?? 0) };
}

function reference(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`;
}
import { XlsParseError } from "../errors";
