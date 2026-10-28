import {
  type FormulaAst,
  type FormulaBinaryOperator,
  type FormulaCallAst,
  type FormulaCellCoordinate,
  type FormulaDialect,
  type FormulaMissingAst,
  type FormulaParseOptions,
  type FormulaRangeEdge,
  type FormulaReferenceAst,
  type FormulaReferenceInfo,
  type FormulaReferenceArea,
  type FormulaWorkbookSource,
  type FormulaTableColumnSelector,
} from "./ast";
import { type FormulaErrorCode, FormulaParseError } from "./errors";
import type { FormulaToken } from "./lexer";
import { tokenizeFormula } from "./lexer";
import { isWholeColumn, isWholeRow, maximumColumns, maximumRows } from "./references";
import { isFormulaErrorToken } from "./value";

const operatorPrecedence: Partial<Record<FormulaBinaryOperator, number>> = {
  "=": 10,
  ",": 5,
  "<>": 10,
  "<": 10,
  "<=": 10,
  ">": 10,
  ">=": 10,
  "&": 20,
  "+": 30,
  "-": 30,
  "*": 40,
  "/": 40,
  "^": 50,
  ":": 100,
  "!": 100,
  " ": 100,
  "~": 100,
};

class FormulaParser {
  readonly source: string;
  readonly dialect: FormulaDialect;
  readonly tokens: FormulaToken[];
  private index = 0;

  constructor(source: string, dialect: FormulaDialect) {
    this.source = source;
    this.dialect = dialect;
    try {
      this.tokens = tokenizeFormula(source, dialect);
    } catch (error) {
      const diagnostic = error as {
        code?: FormulaErrorCode;
        message?: string;
        start?: number;
        end?: number;
      };
      throw new FormulaParseError(
        {
          code: diagnostic.code ?? "unsupported-syntax",
          message: diagnostic.message ?? "invalid formula",
          start: diagnostic.start ?? 0,
          end: diagnostic.end ?? source.length,
          dialect,
        },
        { cause: error },
      );
    }
  }

  parse(): FormulaAst {
    if (this.peek().kind === "end") return this.fail("empty-formula", "formula is empty", 0, 0);
    const hasPrefix =
      this.peekRaw().kind === "operator" && this.peekRaw().text === "=" && this.index === 0;
    if (hasPrefix) {
      this.consume();
      if (this.peek().kind === "end") return this.fail("empty-formula", "formula is empty", 0, 1);
    }
    const ast = this.parseExpression(0);
    this.skipWhitespace();
    if (this.peek().kind !== "end") {
      const token = this.peek();
      this.fail("expected-end", `unexpected "${token.text}"`, token.start, token.end);
    }
    return {
      ...ast,
      ...(hasPrefix ? { prefix: "=" as const } : {}),
      raw: this.source,
      start: 0,
      end: this.source.length,
      dialect: this.dialect,
    };
  }

  private fail(code: FormulaErrorCode, message: string, start: number, end: number): never {
    throw new FormulaParseError({ code, message, start, end, dialect: this.dialect });
  }

  private peek(offset = 0): FormulaToken {
    let index = this.index;
    for (let seen = 0; seen < offset; seen += 1) {
      index += 1;
      while (this.tokens[index]?.kind === "whitespace") index += 1;
    }
    return this.tokens[index] ?? this.tokens[this.tokens.length - 1]!;
  }

  private skipWhitespace(): boolean {
    let found = false;
    while (this.peekRaw().kind === "whitespace") {
      found = true;
      this.index += 1;
    }
    return found;
  }

  private peekRaw(): FormulaToken {
    return this.tokens[this.index] ?? this.tokens[this.tokens.length - 1]!;
  }

  private consume(): FormulaToken {
    const token = this.peekRaw();
    if (token.kind !== "end") this.index += 1;
    return token;
  }

  private consumeThrough(end: number): void {
    while (this.peekRaw().kind !== "end" && this.peekRaw().start < end) this.consume();
  }

  private parseExpression(minimumPrecedence = 0, stopAt: Set<string> = new Set()): FormulaAst {
    let left = this.parsePrimary();
    for (;;) {
      const hadWhitespace = this.skipWhitespace();
      const token = this.peek();
      if (stopAt.has(token.kind) && token.kind !== "operator") break;
      if (token.kind === "spill") {
        this.consume();
        if (left.kind !== "reference" || !left.reference)
          this.fail("expected-expression", "# requires a reference", token.start, token.end);
        if (left.reference.table)
          this.fail(
            "invalid-table-reference",
            "structured references do not support spill operator",
            token.start,
            token.end,
          );
        if (
          left.reference.sheets?.range ||
          (left.reference.sheets?.names.length ?? 0) > 1 ||
          left.reference.areas?.some((area) => area.toSheet !== undefined)
        )
          this.fail(
            "invalid-reference",
            "3D references do not support spill operator",
            token.start,
            token.end,
          );
        left = {
          ...left,
          reference: { ...left.reference, spill: true },
          raw: this.source.slice(left.start, token.end),
          end: token.end,
        };
        continue;
      }
      if (token.kind === "operator" && token.text === "%") {
        const percent = this.consume();
        left = {
          kind: "postfix",
          operator: "%",
          operand: left,
          raw: this.source.slice(left.start, percent.end),
          start: left.start,
          end: percent.end,
        };
        continue;
      }
      if (token.kind === "open-paren") {
        this.consume();
        const args: FormulaAst[] = [];
        this.skipWhitespace();
        if (this.peek().kind !== "close-paren") {
          for (;;) {
            this.skipWhitespace();
            if (this.peek().kind === "close-paren") break;
            const argument = this.parseExpression(0, new Set(["comma", "close-paren"]));
            args.push(argument);
            this.skipWhitespace();
            if (this.peek().kind === "comma") {
              this.consume();
              continue;
            }
            break;
          }
        }
        this.skipWhitespace();
        if (this.peek().kind !== "close-paren") {
          const current = this.peek();
          this.fail("expected-end", "unclosed call application", current.start, current.end);
        }
        const close = this.consume();
        left = {
          kind: "application",
          callee: left,
          args,
          raw: this.source.slice(left.start, close.end),
          start: left.start,
          end: close.end,
        };
        continue;
      }
      if (hadWhitespace && this.startsOperand(token)) {
        const operator: FormulaBinaryOperator = " ";
        const precedence = operatorPrecedence[operator]!;
        if (precedence < minimumPrecedence) break;
        const right = this.parseExpression(precedence + 1, stopAt);
        left = this.binary(left, right, operator);
        continue;
      }
      if (token.kind === "comma" && stopAt.has("comma")) break;
      if (token.kind === "comma") {
        this.consume();
        const right = this.parseExpression(6, stopAt);
        left = this.binary(left, right, ",");
        continue;
      }
      if (token.kind !== "operator") break;
      if (!isFormulaBinaryOperator(token.text)) break;
      const operator: FormulaBinaryOperator = token.text;
      if (token.text === "," && stopAt.has("comma")) break;
      const precedence = operatorPrecedence[operator];
      if (precedence === undefined || precedence < minimumPrecedence) break;
      this.consume();
      const right = this.parseExpression(precedence + 1, stopAt);
      left = this.binary(left, right, operator);
    }
    return left;
  }

  private startsOperand(token: FormulaToken): boolean {
    return (
      ["number", "string", "boolean", "error", "identifier", "open-paren", "open-brace"].includes(
        token.kind,
      ) ||
      (token.kind === "operator" && "+-@$[".includes(token.text))
    );
  }

  private binary(left: FormulaAst, right: FormulaAst, operator: FormulaBinaryOperator): FormulaAst {
    return {
      kind: "binary",
      operator,
      left,
      right,
      raw: this.source.slice(left.start, right.end),
      start: left.start,
      end: right.end,
    };
  }

  private parsePrimary(): FormulaAst {
    this.skipWhitespace();
    const token = this.peekRaw();
    const makeReference = (end: number, reference?: FormulaReferenceInfo): FormulaReferenceAst => ({
      kind: "reference",
      raw: this.source.slice(token.start, end),
      start: token.start,
      end,
      reference,
    });

    if (token.kind === "number") {
      if (this.source[token.end] === ":") {
        const wholeRowRange = this.parseCellRange(token.start);
        if (
          wholeRowRange?.startEdge !== undefined &&
          isWholeRow(wholeRowRange.startEdge) &&
          wholeRowRange.endEdge !== undefined &&
          isWholeRow(wholeRowRange.endEdge)
        ) {
          this.consumeThrough(wholeRowRange.end);
          return makeReference(wholeRowRange.end, rangeInfo([], wholeRowRange));
        }
      }
      if (typeof token.value !== "number")
        this.fail("expected-expression", "invalid number literal", token.start, token.end);
      this.consume();
      return {
        kind: "number",
        raw: token.text,
        start: token.start,
        end: token.end,
        value: token.value,
      };
    }
    if (token.kind === "string") {
      if (typeof token.value !== "string")
        this.fail("expected-expression", "invalid string literal", token.start, token.end);
      this.consume();
      return {
        kind: "string",
        raw: token.text,
        start: token.start,
        end: token.end,
        value: token.value,
      };
    }
    if (token.kind === "boolean" && this.source.slice(token.end).trimStart().startsWith("(")) {
      return this.parseCall();
    }
    if (token.kind === "boolean") {
      if (typeof token.value !== "boolean")
        this.fail("expected-expression", "invalid boolean literal", token.start, token.end);
      this.consume();
      return {
        kind: "boolean",
        raw: token.text,
        start: token.start,
        end: token.end,
        value: token.value,
      };
    }
    if (token.kind === "error") {
      if (!isFormulaErrorToken(token.value))
        this.fail("expected-expression", "invalid error literal", token.start, token.end);
      this.consume();
      return {
        kind: "error",
        raw: token.text,
        start: token.start,
        end: token.end,
        value: token.value,
      };
    }

    if (token.kind === "open-paren") {
      this.consume();
      const expression = this.parseExpression(0, new Set(["close-paren"]));
      this.skipWhitespace();
      if (this.peek().kind !== "close-paren") {
        const current = this.peek();
        this.fail("expected-expression", "unclosed parenthesis", current.start, current.end);
      }
      const close = this.consume();
      return {
        kind: "group",
        operand: expression,
        raw: this.source.slice(token.start, close.end),
        start: token.start,
        end: close.end,
      };
    }

    if (token.kind === "open-brace") {
      return this.parseArray();
    }

    if (token.kind === "operator" && (token.text === "+" || token.text === "-")) {
      this.consume();
      const operand = this.parsePrimary();
      return this.parsePowerOperand({
        kind: "unary",
        operator: token.text === "-" ? "-" : "+",
        operand,
        raw: this.source.slice(token.start, operand.end),
        start: token.start,
        end: operand.end,
      });
    }

    if (token.kind === "operator" && token.text === "@") {
      this.consume();
      const operand = this.parsePrimary();
      if (operand.kind === "name") {
        return {
          ...operand,
          implicitIntersection: true,
          raw: this.source.slice(token.start, operand.end),
        };
      }
      if (operand.kind === "reference") {
        return {
          ...operand,
          reference: { ...operand.reference, implicitIntersection: true },
          raw: this.source.slice(token.start, operand.end),
        };
      }
      this.fail("expected-expression", "@ requires a reference", token.start, operand.end);
    }

    if (token.kind === "operator" && (token.text === "$" || token.text === "[")) {
      const reference = this.parseReference(token.start);
      if (!reference) this.fail("invalid-reference", "invalid reference", token.start, token.end);
      this.consumeThrough(reference.end);
      return makeReference(reference.end, reference.info);
    }

    if (token.kind === "operator" && token.text === "'") {
      const reference = this.parseReference(token.start);
      if (!reference)
        this.fail("invalid-reference", "invalid quoted sheet reference", token.start, token.end);
      this.consumeThrough(reference.end);
      return makeReference(reference.end, reference.info);
    }

    if (token.kind === "identifier") {
      const next = this.peek(1);
      if (next.kind === "open-paren") return this.parseCall();
      if (next.kind === "operator" && next.text === "[") return this.parseTableReference();
      const reference = this.parseReference(token.start);
      if (reference) {
        this.consumeThrough(reference.end);
        return makeReference(reference.end, reference.info);
      }
      if (this.looksLikeExcelCell(token.start))
        this.fail(
          "invalid-reference",
          "cell reference exceeds Excel worksheet boundaries",
          token.start,
          token.end,
        );
      this.consume();
      return {
        kind: "name",
        name: typeof token.value === "string" ? token.value : token.text,
        raw: token.text,
        start: token.start,
        end: token.end,
      };
    }

    this.fail("expected-expression", `unexpected "${token.text}"`, token.start, token.end);
  }

  private parsePowerOperand(left: FormulaAst): FormulaAst {
    let result = left;
    while (this.peekRaw().kind === "operator" && this.peekRaw().text === "^") {
      this.consume();
      const right = this.parsePrimary();
      result = this.binary(result, right, "^");
    }
    return result;
  }

  private parseArray(): FormulaAst {
    const open = this.consume();
    const rows: FormulaAst[][] = [];
    let row: FormulaAst[] = [];
    for (;;) {
      this.skipWhitespace();
      if (this.peek().kind === "close-brace") {
        const close = this.consume();
        if (row.length > 0) rows.push(row);
        const firstWidth = rows[0]?.length ?? 0;
        if (rows.some((r) => r.length !== firstWidth)) {
          throw new FormulaParseError({
            code: "unsupported-syntax",
            message: "array literal rows must have equal column count",
            start: open.start,
            end: close.end,
            dialect: this.dialect,
          });
        }
        return {
          kind: "array",
          rows,
          raw: this.source.slice(open.start, close.end),
          start: open.start,
          end: close.end,
        };
      }
      const value = this.parseExpression(0, new Set(["comma", "semicolon", "operator"]));
      row.push(value);
      this.skipWhitespace();
      const delimiter = this.peek();
      if (
        delimiter.kind === "comma" ||
        delimiter.kind === "semicolon" ||
        (delimiter.kind === "operator" && delimiter.text === "|")
      ) {
        this.consume();
        if (this.startsRowDelimiter(delimiter)) {
          rows.push(row);
          row = [];
        }
        continue;
      }
      if (delimiter.kind === "close-brace") continue;
      this.fail("expected-expression", "invalid array literal", delimiter.start, delimiter.end);
    }
  }

  private startsRowDelimiter(token: FormulaToken): boolean {
    if (this.dialect === "openformula") return token.kind === "operator" && token.text === "|";
    return token.kind === "semicolon";
  }

  private parseCall(): FormulaAst {
    const nameToken = this.consume();
    this.skipWhitespace();
    this.consume();
    const args: FormulaAst[] = [];
    const separators: FormulaToken[] = [];
    const allowSemicolon = this.dialect === "openformula";
    for (;;) {
      this.skipWhitespace();
      const current = this.peek();
      if (current.kind === "close-paren") break;
      if (current.kind === "comma" || (current.kind === "semicolon" && allowSemicolon)) {
        args.push(this.missing(current.start, current.start));
        separators.push(this.consume());
        continue;
      }
      args.push(this.parseExpression(0, new Set(["comma", "semicolon"])));
      this.skipWhitespace();
      const delimiter = this.peek();
      if (delimiter.kind === "comma" || (delimiter.kind === "semicolon" && allowSemicolon)) {
        separators.push(this.consume());
        continue;
      }
      break;
    }
    this.skipWhitespace();
    const close = this.peek();
    if (close.kind !== "close-paren") {
      this.fail("expected-expression", "unclosed function call", close.start, close.end);
    }
    const closing = this.consume();
    while (separators.length >= args.length && separators.length > 0) {
      args.push(this.missing(closing.start, closing.start));
    }
    const ast: FormulaCallAst = {
      kind: "call",
      name: typeof nameToken.value === "string" ? nameToken.value : nameToken.text,
      args,
      raw: this.source.slice(nameToken.start, closing.end),
      start: nameToken.start,
      end: closing.end,
    };
    return ast;
  }

  private missing(start: number, end: number): FormulaMissingAst {
    return { kind: "missing", raw: "", start, end };
  }

  private parseTableReference(): FormulaAst {
    const name = this.consume();
    const open = this.consume();
    let depth = 1;
    let index = open.end;
    while (index < this.source.length) {
      const character = this.source[index]!;
      if (character === "'" && isTableSpecialCharacter(this.source[index + 1] ?? "")) {
        index += 2;
        continue;
      }
      if (character === "[") depth += 1;
      if (character === "]") {
        depth -= 1;
        if (depth === 0) break;
      }
      index += 1;
    }
    if (depth !== 0)
      this.fail(
        "invalid-table-reference",
        "unclosed table reference",
        open.start,
        this.source.length,
      );
    this.consumeThrough(index + 1);
    const raw = this.source.slice(name.start, index + 1);
    const opening = raw.indexOf("[");
    const closing = raw.lastIndexOf("]");
    const specifier = raw.slice(opening, closing + 1);
    const selector = parseTableSpecifier(specifier);
    if (!selector)
      this.fail("invalid-table-reference", "invalid table reference", name.start, index + 1);
    return {
      kind: "reference",
      raw,
      start: name.start,
      end: index + 1,
      reference: {
        table: {
          table: typeof name.value === "string" ? name.value : name.text,
          ...(selector?.item ? { item: selector.item } : {}),
          selectors: selector?.selectors ?? [],
        },
      },
    };
  }

  private parseReference(start: number): { end: number; info?: FormulaReferenceInfo } | undefined {
    if (this.source[start] === "[") return this.parseBracketedWorkbookReference(start);
    if (this.source[start] === "'") return this.parseQuotedSheetReference(start);
    return this.parseUnquotedSheetReference(start);
  }

  private parseBracketedWorkbookReference(
    start: number,
  ): { end: number; info?: FormulaReferenceInfo } | undefined {
    if (this.dialect === "excel") {
      const workbook = this.parseExcelWorkbookSource(start);
      if (!workbook) return undefined;
      const reference = this.parseUnquotedSheetReference(workbook.end);
      if (!reference?.info) return undefined;
      return {
        end: reference.end,
        info: {
          workbook: workbook.workbook,
          ...reference.info,
          areas: reference.info.areas?.map((area) => ({ ...area, workbook: workbook.workbook })),
        },
      };
    }
    let index = start + 1;
    let depth = 1;
    while (index < this.source.length) {
      const character = this.source[index]!;
      if (character === `'`) {
        index += 1;
        while (index < this.source.length) {
          if (this.source[index] === `'` && this.source[index + 1] === `'`) {
            index += 2;
            continue;
          }
          if (this.source[index] === `'`) break;
          index += 1;
        }
      } else if (character === "[") {
        depth += 1;
      } else if (character === "]") {
        depth -= 1;
        if (depth === 0) {
          const content = this.source.slice(start + 1, index);
          const info = parseOpenFormulaReference(content);
          return info ? { end: index + 1, info } : undefined;
        }
      }
      index += 1;
    }
    return undefined;
  }

  private parseQuotedSheetReference(
    start: number,
  ): { end: number; info?: FormulaReferenceInfo } | undefined {
    const firstName = this.parseQuotedSheetName(start);
    if (!firstName) return undefined;
    let index = firstName.end;
    if (this.source[index] === ":" && this.source[index + 1] === "'") {
      const secondName = this.parseQuotedSheetName(index + 1);
      if (!secondName) return undefined;
      index = secondName.end;
      if (this.source[index] !== `!`) return undefined;
      const range = this.parseCellRange(index + 1);
      if (!range) return undefined;
      const names = [firstName.name, secondName.name];
      return {
        end: range.end,
        info: {
          sheets: { names, range: true },
          ...rangeInfo(names, range),
        },
      };
    }
    if (this.source[index] !== `!`) return undefined;
    index += 1;
    const range = this.parseCellRange(index);
    if (!range) return undefined;
    return {
      end: range.end,
      info: {
        sheets: { names: [firstName.name], range: false },
        ...rangeInfo([firstName.name], range),
      },
    };
  }

  private parseQuotedSheetName(start: number): { name: string; end: number } | undefined {
    let index = start + 1;
    let name = "";
    while (index < this.source.length) {
      if (this.source[index] === `'` && this.source[index + 1] === `'`) {
        name += `'`;
        index += 2;
        continue;
      }
      if (this.source[index] === `'`) break;
      name += this.source[index]!;
      index += 1;
    }
    if (this.source[index] !== `'`) return undefined;
    return { name, end: index + 1 };
  }

  private parseUnquotedSheetReference(
    start: number,
  ): { end: number; info?: FormulaReferenceInfo } | undefined {
    let index = start;
    while (index < this.source.length && !`![]();,"+*/^&=<>@%~-`.includes(this.source[index]!))
      index += 1;
    if (this.source[index] === `!`) {
      const segment = this.source.slice(start, index);
      if (!segment) return undefined;
      const names = segment
        .split(":")
        .map((name) => name.trim().replace(/^'/u, "").replace(/'$/u, ""));
      const firstEdge = this.parseCellEdge(index + 1);
      if (
        firstEdge &&
        this.source[firstEdge.end] === ":" &&
        this.parseUnquotedSheetReference(firstEdge.end + 1)?.info?.areas?.[0]?.fromSheet
      ) {
        const second = this.parseUnquotedSheetReference(firstEdge.end + 1);
        const secondInfo = second?.info;
        const secondArea = secondInfo?.areas?.[0];
        const secondEnd = secondArea?.end;
        if (
          secondInfo?.sheets?.names.length === 1 &&
          secondArea?.fromSheet &&
          secondEnd &&
          referenceAxisMatches(firstEdge, secondEnd)
        ) {
          const crossNames = [names[0]!, secondInfo.sheets.names[0]!];
          return {
            end: second!.end,
            info: {
              sheets: { names: crossNames, range: true },
              ...rangeInfo(crossNames, {
                ...(firstEdge.kind === "cell"
                  ? { start: firstEdge.coordinate, startEdge: firstEdge.coordinate }
                  : { startEdge: normalizeRangeEdge(firstEdge) }),
                ...(isCellEdge(secondEnd)
                  ? { endCoordinate: secondEnd, endEdge: secondEnd }
                  : { endEdge: secondEnd }),
                end: second!.end,
              }),
            },
          };
        }
      }
      const range = this.parseCellRange(index + 1);
      if (!range) return undefined;
      if (this.source[range.end] === ":" && names.length === 1) {
        const second = this.parseUnquotedSheetReference(range.end + 1);
        const secondInfo = second?.info;
        const secondArea = secondInfo?.areas?.[0];
        const secondEnd = secondArea?.end;
        if (
          secondInfo?.sheets?.names.length === 1 &&
          secondArea?.fromSheet &&
          range.startEdge &&
          secondEnd &&
          referenceAxisMatches(range.startEdge, secondEnd)
        ) {
          const crossNames = [names[0]!, secondInfo.sheets.names[0]!];
          return {
            end: second!.end,
            info: {
              sheets: { names: crossNames, range: true },
              ...rangeInfo(crossNames, {
                start: range.start,
                startEdge: range.startEdge,
                ...(isCellEdge(secondEnd)
                  ? { endCoordinate: secondEnd, endEdge: secondEnd }
                  : { endEdge: secondEnd }),
                end: second!.end,
              }),
            },
          };
        }
      }
      return {
        end: range.end,
        info: {
          sheets: { names, range: names.length > 1 },
          ...rangeInfo(names, range),
        },
      };
    }
    const hasRangeSeparator = this.source.slice(start, index).includes(":");
    const localRange = this.parseCellRange(start);
    if (localRange && localRange.start === undefined && !hasRangeSeparator) return undefined;
    if (localRange) {
      const next = localRange.end;
      if (this.source[next] === ":") {
        const second = this.parseUnquotedSheetReference(next + 1);
        const secondInfo = second?.info;
        const secondArea = secondInfo?.areas?.[0];
        const secondEnd = secondArea?.end;
        const localSheet = this.unquotedSheetNameBefore(start);
        if (
          secondInfo?.sheets?.names.length === 1 &&
          localSheet &&
          secondArea?.fromSheet &&
          localRange.startEdge &&
          secondEnd &&
          referenceAxisMatches(localRange.startEdge, secondEnd)
        ) {
          const names = [localSheet, secondInfo!.sheets!.names[0]!];
          return {
            end: second!.end,
            info: {
              sheets: { names, range: true },
              ...rangeInfo(names, {
                start: localRange.start,
                startEdge: localRange.startEdge,
                ...(isCellEdge(secondEnd)
                  ? { endCoordinate: secondEnd, endEdge: secondEnd }
                  : { endEdge: secondEnd }),
                end: second!.end,
              }),
            },
          };
        }
      }
      const spill = this.source[next] === "#";
      const info = rangeInfo([], localRange);
      return {
        end: spill ? next + 1 : next,
        info: { ...info, spill: spill || undefined },
      };
    }
    return undefined;
  }

  private unquotedSheetNameBefore(position: number): string | undefined {
    let start = position;
    const delimiters = `![]();,"+*/^&=<>@%~-`;
    while (start > 0 && !delimiters.includes(this.source[start - 1]!)) start -= 1;
    if (start === position || this.source[start - 1] !== "!") return undefined;
    return this.source.slice(start, position).replace(/^'/u, "").replace(/'$/u, "");
  }

  private parseExcelWorkbookSource(
    start: number,
  ): { workbook: FormulaWorkbookSource; end: number } | undefined {
    let index = start + 1;
    let identity: string;
    if (this.source[index] === "'") {
      index += 1;
      identity = "";
      while (index < this.source.length) {
        if (this.source[index] === "'" && this.source[index + 1] === "'") {
          identity += "'";
          index += 2;
          continue;
        }
        if (this.source[index] === "'") break;
        identity += this.source[index];
        index += 1;
      }
      if (this.source[index] !== "'" || this.source[index + 1] !== "]") return undefined;
      index += 2;
    } else {
      const close = this.source.indexOf("]", index);
      if (close === -1) return undefined;
      identity = this.source.slice(index, close);
      index = close + 1;
    }
    if (!identity) return undefined;
    const workbook = /^\d+$/u.test(identity)
      ? { index: Number(identity) }
      : /[/\\]/u.test(identity)
        ? { path: identity }
        : { name: identity };
    return { workbook, end: index };
  }

  private parseCellRange(start: number):
    | {
        start?: FormulaCellCoordinate;
        endCoordinate?: FormulaCellCoordinate;
        startEdge?: FormulaRangeEdge;
        endEdge?: FormulaRangeEdge;
        end: number;
      }
    | undefined {
    const first = this.parseCellEdge(start);
    if (!first) return undefined;
    let firstEdge = first;
    let lastEdge = first;
    let cursor = first.end;
    while (this.source[cursor] === ":") {
      const next = this.parseCellEdge(cursor + 1);
      if (!next) {
        if (this.looksLikeExcelReference(cursor + 1)) {
          this.fail(
            "invalid-reference",
            "range edge exceeds Excel worksheet boundaries",
            cursor,
            this.source.length,
          );
        }
        return undefined;
      }
      if (firstEdge.kind !== next.kind) this.failInvalidRangeAxis(cursor + 1);
      if (
        firstEdge.kind === "edge" &&
        next.kind === "edge" &&
        (typeof firstEdge.columns === "number") !== (typeof next.columns === "number")
      ) {
        this.failInvalidRangeAxis(cursor + 1);
      }
      firstEdge = smallerRangeEdge(firstEdge, next);
      lastEdge = largerRangeEdge(lastEdge, next);
      cursor = next.end;
    }
    if (firstEdge.kind === "cell" && lastEdge.kind === "cell") {
      return {
        start: firstEdge.coordinate,
        endCoordinate: lastEdge.coordinate,
        startEdge: firstEdge.coordinate,
        endEdge: lastEdge.coordinate,
        end: cursor,
      };
    }
    return {
      startEdge: normalizeRangeEdge(firstEdge),
      endEdge: normalizeRangeEdge(lastEdge),
      end: cursor,
    };
  }

  private failInvalidRangeAxis(start: number): never {
    this.fail("invalid-reference", "range edges use different reference axes", start, start + 1);
  }

  private looksLikeExcelReference(start: number): boolean {
    const match = /^\$?(?:[A-Za-z]{1,3}\$?\d+|[A-Za-z]{1,3}|\d+)/u.exec(this.source.slice(start));
    if (!match) return false;
    const end = start + match[0].length;
    return /[\s,;)\]!:#]/u.test(this.source[end] ?? "") === true || end === this.source.length;
  }

  private looksLikeExcelCell(start: number): boolean {
    const match = /^\$?[A-Za-z]{1,3}\$?\d+/u.exec(this.source.slice(start));
    if (!match) return false;
    const end = start + match[0].length;
    return (
      /[\s,;)\]!:#+*/^&=<>@%-]/u.test(this.source[end] ?? "") === true || end === this.source.length
    );
  }

  private parseCellEdge(start: number): FormulaCellEdgeResult | undefined {
    let index = start;
    let absoluteColumn = false;
    let absoluteRow = false;
    if (this.source[index] === "$" && /[A-Za-z]/u.test(this.source[index + 1] ?? "")) {
      absoluteColumn = true;
      index += 1;
    }
    let column = 0;
    let columns = 0;
    while (index < this.source.length && /[A-Za-z]/u.test(this.source[index]!)) {
      column = column * 26 + (this.source[index]!.toUpperCase().charCodeAt(0) - 64);
      columns += 1;
      index += 1;
    }
    if (columns > 3) return undefined;
    if (
      this.source[index] === "$" &&
      this.source[index + 1]! >= "0" &&
      this.source[index + 1]! <= "9"
    ) {
      absoluteRow = true;
      index += 1;
    }
    let row = 0;
    let rows = 0;
    while (index < this.source.length && this.source[index]! >= "0" && this.source[index]! <= "9") {
      row = row * 10 + Number(this.source[index]!);
      rows += 1;
      index += 1;
    }
    if (columns > 0 && rows > 0 && /[A-Za-z]/u.test(this.source[index] ?? "")) return undefined;
    if (columns > 0 && rows > 0) {
      if (column > maximumColumns || row > maximumRows) return undefined;
      return { kind: "cell", coordinate: { row, column, absoluteRow, absoluteColumn }, end: index };
    }
    if (columns > 0 && rows === 0)
      return column <= maximumColumns
        ? { kind: "edge", end: index, column, columns, absoluteColumn }
        : undefined;
    if (columns === 0 && rows > 0)
      return row <= maximumRows ? { kind: "edge", end: index, row, rows, absoluteRow } : undefined;
    return undefined;
  }
}

export function parseFormula(formula: string, options: FormulaParseOptions = {}): FormulaAst {
  try {
    return new FormulaParser(formula, options.dialect ?? "excel").parse();
  } catch (error) {
    if (!(error instanceof RangeError)) throw error;
    throw new FormulaParseError(
      {
        code: "unsupported-syntax",
        message: "formula nesting is too deep",
        start: 0,
        end: formula.length,
        dialect: options.dialect ?? "excel",
      },
      { cause: error },
    );
  }
}

type CellRangeResult = NonNullable<
  ReturnType<InstanceType<typeof FormulaParser>["parseCellRange"]>
>;

function normalizeRangeEdge(edge: FormulaCellEdgeResult): FormulaRangeEdge {
  if (edge.kind === "cell") return edge.coordinate;
  if (typeof edge.columns === "number")
    return {
      kind: "whole-column",
      column: edge.column ?? 0,
      absolute: edge.absoluteColumn === true,
    };
  return {
    kind: "whole-row",
    row: edge.row ?? 0,
    absolute: edge.absoluteRow === true,
  };
}

function smallerRangeEdge(left: FormulaCellEdgeResult, right: FormulaCellEdgeResult) {
  return compareRangeEdges(left, right) <= 0 ? left : right;
}

function largerRangeEdge(left: FormulaCellEdgeResult, right: FormulaCellEdgeResult) {
  return compareRangeEdges(left, right) >= 0 ? left : right;
}

function compareRangeEdges(left: FormulaCellEdgeResult, right: FormulaCellEdgeResult): number {
  if (left.kind === "cell" && right.kind === "cell") {
    return (
      left.coordinate.row - right.coordinate.row || left.coordinate.column - right.coordinate.column
    );
  }
  if (left.kind === "edge" && right.kind === "edge") {
    if (left.columns !== undefined && right.columns !== undefined) {
      return (left.column ?? 0) - (right.column ?? 0);
    }
    return (left.row ?? 0) - (right.row ?? 0);
  }
  return 0;
}

function isFormulaBinaryOperator(value: string): value is FormulaBinaryOperator {
  return value in operatorPrecedence;
}

function referenceAxisMatches(
  left: FormulaCellEdgeResult | FormulaRangeEdge,
  right: FormulaRangeEdge,
): boolean {
  if (!("kind" in left)) return !isWholeColumn(right) && !isWholeRow(right);
  if (left.kind === "cell" && "coordinate" in left)
    return !isWholeColumn(right) && !isWholeRow(right);
  if (left.kind === "whole-column") return isWholeColumn(right);
  return isWholeRow(right);
}

type FormulaCellEdgeResult =
  | { kind: "cell"; coordinate: FormulaCellCoordinate; end: number }
  | {
      kind: "edge";
      end: number;
      column?: number;
      columns?: number;
      absoluteColumn?: boolean;
      row?: number;
      rows?: number;
      absoluteRow?: boolean;
    };

function rangeInfo(
  sheets: string[],
  range: CellRangeResult,
): Pick<FormulaReferenceInfo, "start" | "end" | "areas"> {
  const start = range.startEdge ?? range.start;
  const end = range.endEdge ?? range.endCoordinate ?? start;
  if (!start || !end) return { start: range.start, end: range.endCoordinate };
  const area: FormulaReferenceArea = {
    start,
    end,
    ...(sheets.length === 2 ? { fromSheet: sheets[0], toSheet: sheets[1] } : {}),
    ...(sheets.length === 1 ? { fromSheet: sheets[0] } : {}),
  };
  return { start: range.start, end: range.endCoordinate, areas: [area] };
}

function parseOpenFormulaReference(content: string): FormulaReferenceInfo | undefined {
  const edges = splitOpenFormulaEdges(content).map(parseOpenFormulaEdge);
  if (!edges.length || edges.some((edge) => !edge)) return undefined;
  if (edges.length > 2) return undefined;
  const start = edges[0]!;
  const end = edges[1] ?? start;
  const workbooks = edges.filter(
    (edge): edge is OpenFormulaEdge & { workbook: FormulaWorkbookSource } =>
      edge?.workbook !== undefined,
  );
  const workbook = workbooks[0]?.workbook;
  if (
    workbook &&
    workbooks.some(
      (edge) => openFormulaWorkbookKey(edge.workbook) !== openFormulaWorkbookKey(workbook),
    )
  )
    return undefined;
  const sheets = edges.filter((edge): edge is OpenFormulaEdge => edge?.sheet !== undefined);
  const names = [...new Set(sheets.map((edge) => edge.sheet!))];
  if (referenceAxis(edges[0]!.edge) !== referenceAxis(edges[edges.length - 1]!.edge)) {
    return undefined;
  }
  const area: FormulaReferenceArea = {
    ...(workbook ? { workbook } : {}),
    start: start.edge,
    ...(end !== start ? { end: end.edge } : {}),
    ...(start.sheet ? { fromSheet: start.sheet } : {}),
    ...(end.sheet && end.sheet !== start.sheet ? { toSheet: end.sheet } : {}),
  };
  return {
    ...(names.length ? { sheets: { names, range: names.length > 1 } } : {}),
    areas: [area],
    ...(isCellEdge(start.edge) ? { start: start.edge } : {}),
    ...(isCellEdge(end.edge) ? { end: end.edge } : {}),
  };
}

function parseTableSpecifier(value: string):
  | {
      item?: "all" | "data" | "headers" | "totals" | "this-row";
      selectors: FormulaTableColumnSelector[];
    }
  | undefined {
  if (!value.startsWith("[") || !value.endsWith("]")) return undefined;
  const inner = tableSpecifierContent(value);
  if (inner === undefined) return undefined;
  const itemSource = /^\s*\[\s*#/u.test(inner) ? inner.slice(1) : inner;
  const itemMatch = /^\s*(#All|#Data|#Headers|#Totals|#This Row)\]?\s*(?:,|$)/iy.exec(itemSource);
  const item = itemMatch
    ? (itemMatch[1]!.toLowerCase().replace("#", "").replace(" ", "-") as
        | "all"
        | "data"
        | "headers"
        | "totals"
        | "this-row")
    : undefined;
  const remainder = itemMatch ? itemSource.slice(itemMatch[0].length) : inner;
  const thisRowMatch = /^\s*@/uy.exec(remainder);
  const normalizedItem = thisRowMatch ? "this-row" : item;
  const selectors = thisRowMatch ? remainder.slice(thisRowMatch[0].length) : remainder;
  if (selectors.trim() === "") {
    if (!normalizedItem) return undefined;
    return { ...(normalizedItem ? { item: normalizedItem } : {}), selectors: [] };
  }
  const selected: FormulaTableColumnSelector[] = [];
  for (const selector of splitTableSelectors(selectors)) {
    if (selector.trim() === "") return undefined;
    if (selector.trim().startsWith("#")) return undefined;
    const range = parseTableColumnRange(selector);
    if (range) {
      selected.push({ kind: "range", start: range[0], end: range[1] });
      continue;
    }
    const column = parseTableColumnSelector(selector);
    if (column === undefined) return undefined;
    selected.push({ kind: "column", column });
  }
  if (!normalizedItem && selected.length === 0) return undefined;
  return { ...(normalizedItem ? { item: normalizedItem } : {}), selectors: selected };
}

function tableSpecifierContent(value: string): string | undefined {
  let content = "";
  let depth = 0;
  for (let index = 1; index < value.length; index += 1) {
    const character = value[index]!;
    if (character === "'" && isTableSpecialCharacter(value[index + 1] ?? "")) {
      content += character + value[index + 1]!;
      index += 1;
      continue;
    }
    if (character === "[") depth += 1;
    if (character === "]") {
      if (depth === 0) return index === value.length - 1 ? content : undefined;
      depth -= 1;
    }
    content += character;
  }
  return undefined;
}

function splitTableSelectors(value: string): string[] {
  const selectors: string[] = [];
  let start = 0;
  let bracketDepth = 0;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]!;
    if (character === "'" && isTableSpecialCharacter(value[index + 1] ?? "")) {
      index += 1;
      continue;
    }
    if (character === "[") bracketDepth += 1;
    if (character === "]") bracketDepth -= 1;
    if (character === "," && bracketDepth === 0) {
      selectors.push(value.slice(start, index));
      start = index + 1;
    }
  }
  selectors.push(value.slice(start));
  return selectors;
}

function isTableSpecialCharacter(character: string): boolean {
  return ["#", "[", "]", "'"].includes(character);
}

function parseTableColumnSelector(value: string): string | undefined {
  if (!value.startsWith("[")) {
    const source = value.trim();
    if (source === "") return undefined;
    let column = "";
    let index = 0;
    while (index < source.length) {
      const character = source[index]!;
      if (character === "'" && isTableSpecialCharacter(source[index + 1] ?? "")) {
        column += source[index + 1]!;
        index += 2;
        continue;
      }
      if (character === "'") return undefined;
      if (character === "[" || character === "]") return undefined;
      column += character;
      index += 1;
    }
    return column;
  }
  let index = 1;
  let column = "";
  while (index < value.length) {
    const character = value[index]!;
    if (character === "'" && isTableSpecialCharacter(value[index + 1] ?? "")) {
      column += value[index + 1]!;
      index += 2;
      continue;
    }
    if (character === "'") return undefined;
    if (character === "]") return index === value.length - 1 && column !== "" ? column : undefined;
    column += character;
    index += 1;
  }
  return undefined;
}

function parseTableColumnRange(value: string): [string, string] | undefined {
  if (!value.startsWith("[")) return undefined;
  let index = 1;
  let bracketDepth = 1;
  while (index < value.length) {
    const character = value[index]!;
    if (character === "'" && isTableSpecialCharacter(value[index + 1] ?? "")) {
      index += 2;
      continue;
    }
    if (character === "[") bracketDepth += 1;
    if (character === "]") {
      bracketDepth -= 1;
      if (bracketDepth !== 0) {
        index += 1;
        continue;
      }
      const separator = index + 1;
      if (value[separator] !== ":" || value[separator + 1] !== "[") return undefined;
      const start = parseTableColumnSelector(value.slice(0, separator));
      const end = parseTableColumnSelector(value.slice(separator + 1));
      return start === undefined || end === undefined ? undefined : [start, end];
    }
    index += 1;
  }
  return undefined;
}

interface OpenFormulaEdge {
  workbook?: FormulaWorkbookSource;
  sheet?: string;
  edge: FormulaRangeEdge;
}

function splitOpenFormulaEdges(content: string): string[] {
  const edges: string[] = [];
  let start = 0;
  let quoted = false;
  for (let index = 0; index < content.length; index += 1) {
    if (content[index] === "'") {
      if (content[index + 1] === "'") index += 1;
      else quoted = !quoted;
    }
    if (content[index] === ":" && !quoted) {
      edges.push(content.slice(start, index));
      start = index + 1;
    }
  }
  edges.push(content.slice(start));
  return edges;
}

function parseOpenFormulaWorkbook(
  value: string,
): { workbook: FormulaWorkbookSource; rest: string } | undefined {
  if (value.startsWith("'")) {
    let index = 1;
    let identity = "";
    while (index < value.length) {
      if (value[index] === "'" && value[index + 1] === "'") {
        identity += "'";
        index += 2;
        continue;
      }
      if (value[index] === "'") break;
      identity += value[index];
      index += 1;
    }
    if (value[index] !== "'" || value[index + 1] !== "#") return undefined;
    return {
      workbook: /[/\\]/u.test(identity) ? { path: identity } : { name: identity },
      rest: value.slice(index + 2),
    };
  }
  const separator = value.indexOf("#");
  if (separator === -1) return undefined;
  const identity = value.slice(0, separator);
  if (!identity) return undefined;
  const workbook = /^\d+$/u.test(identity)
    ? { index: Number(identity) }
    : /[/\\]/u.test(identity)
      ? { path: identity }
      : { name: identity };
  return { workbook, rest: value.slice(separator + 1) };
}

function parseOpenFormulaEdge(value: string): OpenFormulaEdge | undefined {
  const workbook = parseOpenFormulaWorkbook(value);
  const normalizedSource = workbook?.rest ?? value;
  const sheet = parseOpenFormulaSheet(normalizedSource);
  if (!sheet) return undefined;
  const referenceSource = sheet?.rest ?? normalizedSource;
  const cell = /^(?:([^.]+)\.)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/u.exec(referenceSource);
  if (cell) {
    return openFormulaEdge(workbook, sheet, {
      row: Number(cell[5]),
      column: parseOpenFormulaColumn(cell[3]!),
      absoluteColumn: cell[2] === "$",
      absoluteRow: cell[4] === "$",
    });
  }
  const column = /^(?:([^.]+)\.)?(\$?)([A-Za-z]{1,3})$/u.exec(referenceSource);
  if (column) {
    return openFormulaEdge(workbook, sheet, {
      kind: "whole-column",
      column: parseOpenFormulaColumn(column[3]!),
      absolute: column[2] === "$",
    });
  }
  const row = /^(?:([^.]+)\.)?(\$?)(\d+)$/u.exec(referenceSource);
  if (row) {
    return openFormulaEdge(workbook, sheet, {
      kind: "whole-row",
      row: Number(row[3]),
      absolute: row[2] === "$",
    });
  }
  return undefined;
}

function openFormulaEdge(
  workbook: { workbook: FormulaWorkbookSource; rest: string } | undefined,
  sheet: { sheet?: string } | undefined,
  edge: FormulaRangeEdge,
): OpenFormulaEdge | undefined {
  if (isCellEdge(edge)) {
    if (edge.row < 1 || edge.row > maximumRows) return undefined;
    if (edge.column < 1 || edge.column > maximumColumns) return undefined;
  } else if (edge.kind === "whole-column" && edge.column > maximumColumns) {
    return undefined;
  } else if (edge.kind === "whole-row" && edge.row > maximumRows) {
    return undefined;
  }
  return {
    ...(workbook ? { workbook: workbook.workbook } : {}),
    ...(sheet ? { sheet: sheet.sheet } : {}),
    edge,
  };
}

function parseOpenFormulaColumn(value: string): number {
  return value
    .split("")
    .reduce((column, character) => column * 26 + character.toUpperCase().charCodeAt(0) - 64, 0);
}

function parseOpenFormulaSheet(value: string): { sheet?: string; rest: string } | undefined {
  if (value.startsWith(".")) return { rest: value.slice(1) };
  const quoted = value.startsWith("'") || value.startsWith("$'");
  if (!quoted) {
    const separator = value.indexOf(".");
    if (separator === -1) return undefined;
    const sheet = value.slice(0, separator).replace(/^\$/u, "");
    if (!sheet || sheet.includes("'")) return undefined;
    return { sheet, rest: value.slice(separator + 1) };
  }
  let index = value.startsWith("$'") ? 2 : 1;
  let sheet = "";
  while (index < value.length) {
    if (value[index] === "'" && value[index + 1] === "'") {
      sheet += "'";
      index += 2;
      continue;
    }
    if (value[index] === "'") break;
    sheet += value[index];
    index += 1;
  }
  if (value[index] !== "'" || value[index + 1] !== ".") return undefined;
  return { sheet, rest: value.slice(index + 2) };
}

function isCellEdge(edge: FormulaRangeEdge): edge is FormulaCellCoordinate {
  return !("kind" in edge);
}

function openFormulaWorkbookKey(workbook: FormulaWorkbookSource): string {
  if (workbook.path !== undefined) return `path:${workbook.path}`;
  if (workbook.index !== undefined) return `index:${workbook.index}`;
  return `name:${workbook.name ?? ""}`;
}

function referenceAxis(edge: FormulaRangeEdge): "cell" | "row" | "column" {
  if (isCellEdge(edge)) return "cell";
  return edge.kind === "whole-row" ? "row" : "column";
}
