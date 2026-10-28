import type { FormulaAst, FormulaDialect } from "./ast";
import { formatReferenceArea } from "./references";

export interface FormulaStringifyOptions {
  dialect?: FormulaDialect;
}

export function stringifyFormula(ast: FormulaAst, options: FormulaStringifyOptions = {}): string {
  const dialect = options.dialect ?? ast.dialect ?? "excel";
  const delimiter = dialect === "openformula" ? "|" : ";";
  const columnDelimiter = dialect === "openformula" ? ";" : ",";
  const stringify = (node: FormulaAst): string => {
    switch (node.kind) {
      case "array":
        return `{${node.rows
          .map((row) => row.map(stringify).join(columnDelimiter))
          .join(delimiter)}}`;
      case "binary":
        return node.operator === " "
          ? `${stringify(node.left)} ${stringify(node.right)}`
          : `${stringify(node.left)}${node.operator}${stringify(node.right)}`;
      case "boolean":
        return node.value ? "TRUE" : "FALSE";
      case "call":
        return `${node.name}(${node.args.map(stringify).join(dialect === "openformula" ? ";" : ",")})`;
      case "application":
        return `${stringify(node.callee)}(${node.args
          .map(stringify)
          .join(dialect === "openformula" ? ";" : ",")})`;
      case "error":
        return node.value;
      case "group":
        return `(${stringify(node.operand)})`;
      case "missing":
        return "";
      case "name":
        return `${node.implicitIntersection ? "@" : ""}${node.name}`;
      case "number":
        return String(node.value);
      case "postfix":
        return `${stringify(node.operand)}%`;
      case "reference":
        const implicitIntersection = node.reference?.implicitIntersection ? "@" : "";
        if (node.reference?.areas?.length)
          return `${implicitIntersection}${node.reference.areas
            .map((area) => formatReferenceArea(area, dialect))
            .join(dialect === "openformula" ? "~" : ",")}${node.reference.spill ? "#" : ""}`;
        if (node.reference?.table) {
          const table = node.reference.table;
          const selectors = [
            ...(table.item
              ? [
                  `[#${table.item.replaceAll("-", " ").replace(/\b[a-z]/gu, (character) => character.toUpperCase())}]`,
                ]
              : []),
            ...table.selectors.map((selector) =>
              selector.kind === "column"
                ? `[${formatTableColumn(selector.column)}]`
                : `[${formatTableColumn(selector.start)}]:[${formatTableColumn(selector.end)}]`,
            ),
          ];
          if (selectors.length === 0) return `${implicitIntersection}${table.table}`;
          if (!table.item && table.selectors.length === 1 && table.selectors[0]?.kind === "column")
            return `${implicitIntersection}${table.table}${selectors[0]}`;
          return `${implicitIntersection}${table.table}[${selectors.join(",")}]`;
        }
        return `${implicitIntersection}${node.raw.replace(/^@/u, "")}${
          node.reference?.spill ? "#" : ""
        }`;
      case "string":
        return `"${node.value.replaceAll('"', '""')}"`;
      case "unary":
        return `${node.operator}${stringify(node.operand)}`;
    }
  };
  return `${ast.prefix ?? ""}${stringify(ast)}`;
}

function formatTableColumn(column: string): string {
  let formatted = "";
  for (const character of column) {
    if (character === "'" || character === "#" || character === "[" || character === "]") {
      formatted += `'`;
    }
    formatted += character;
  }
  return formatted;
}
