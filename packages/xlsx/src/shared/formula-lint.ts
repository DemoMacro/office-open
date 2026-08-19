/**
 * Static formula lint over workbook options.
 *
 * Catches the formula failure classes a JSON schema cannot see before the file
 * reaches Excel: references to sheets the workbook does not contain (Excel
 * evaluates them to #REF!) and unbalanced parentheses or quotes (Excel rejects
 * the formula outright). Best-effort by design — it reads sheet references
 * syntactically and never evaluates anything.
 *
 * @module
 */
import type { WorkbookOptions } from "../parts/file";
import type { WorksheetOptions } from "../parts/worksheet/types";

export interface FormulaIssue {
  /** Location hint, e.g. "Sheet1!B2" or "Sheet1 conditional format A1:A10". */
  location: string;
  formula: string;
  message: string;
}

// \w plus '.' (sheet names like "Q1.Report") and ':' (3-D ranges "A:B!").
const REF_CHARS = /[\w.:]/;

// Cell references ("A1", "$AB$12") swept up while scanning a range segment
// back from '!' (e.g. the "A1:Data" before the second '!' in "Data!A1:Data!B2").
// Skipping them favors misses over false positives — a lint layer must not
// reject a legal formula.
const CELL_REF = /^\$?[A-Za-z]{1,3}\$?\d+$/;

// String literals: an escaped quote inside a literal is written as "".
const STRING_LITERAL = /"(?:[^"]|"")*"/g;

function syntaxIssue(formula: string): string | undefined {
  const stripped = formula.replace(STRING_LITERAL, '""');
  let parens = 0;
  for (const ch of stripped) {
    if (ch === "(") parens++;
    else if (ch === ")") {
      parens--;
      if (parens < 0) return "unbalanced parentheses";
    }
  }
  if (parens !== 0) return "unbalanced parentheses";
  let quotes = 0;
  for (let i = 0; i < formula.length; i++) if (formula.charCodeAt(i) === 0x22) quotes++;
  if (quotes % 2 !== 0) return "unbalanced double quotes";
  return undefined;
}

function sheetReferenceIssues(formula: string, sheetNames: ReadonlySet<string>): string[] {
  // Strip literals first: after this every remaining '!' belongs to a sheet
  // reference — Excel's inequality is '<>' and error literals like #REF! are
  // skipped below, so '!' has no other use in a formula.
  const stripped = formula.replace(STRING_LITERAL, '""');
  const messages: string[] = [];
  for (let i = stripped.indexOf("!"); i !== -1; i = stripped.indexOf("!", i + 1)) {
    const names: string[] = [];
    if (stripped.charCodeAt(i - 1) === 0x27) {
      // Quoted sheet name: 'My Sheet'!A1 — look up no further than its quotes.
      const open = stripped.lastIndexOf("'", i - 2);
      if (open === -1) continue;
      const name = stripped.slice(open + 1, i - 1);
      if (!name) continue;
      // External-workbook ref "[1]'My Sheet'!A1" — target is not this file.
      if (stripped.charCodeAt(open - 1) === 0x5d) continue;
      names.push(name);
    } else {
      let k = i;
      while (k > 0 && REF_CHARS.test(stripped[k - 1]!)) k--;
      const segment = stripped.slice(k, i);
      // Empty (stray '!'), error literal (#REF!), or external ref ([1]Sheet2!).
      if (!segment) continue;
      if (stripped[k - 1] === "#" || stripped[k - 1] === "]") continue;
      for (const name of segment.split(":")) {
        if (!CELL_REF.test(name)) names.push(name);
      }
    }
    for (const name of names) {
      if (!sheetNames.has(name)) {
        messages.push(
          `references sheet "${name}" which does not exist (sheets: ${[...sheetNames].join(", ")})`,
        );
      }
    }
  }
  return [...new Set(messages)];
}

function collectSheetNames(options: WorkbookOptions): Set<string> {
  const names = new Set<string>();
  // Mirror the compiler's default worksheet naming (`Sheet${i+1}`).
  options.worksheets?.forEach((ws, i) => names.add(ws.name ?? `Sheet${i + 1}`));
  // Chartsheet/dialogsheet defaults depend on global sheet order, so only
  // explicit names are lintable — formulas rarely target them anyway.
  for (const sheet of options.chartsheets ?? []) if (sheet.name) names.add(sheet.name);
  for (const sheet of options.dialogsheets ?? []) if (sheet.name) names.add(sheet.name);
  return names;
}

/** Lint all formulas in a workbook: cell formulas, conditional-format rules,
 *  and data-validation formulas. Empty result = no findings. */
export function lintWorkbookFormulas(options: WorkbookOptions): FormulaIssue[] {
  const issues: FormulaIssue[] = [];
  const sheetNames = collectSheetNames(options);

  const check = (formula: string, location: string) => {
    const syntax = syntaxIssue(formula);
    if (syntax) issues.push({ location, formula, message: syntax });
    for (const message of sheetReferenceIssues(formula, sheetNames)) {
      issues.push({ location, formula, message });
    }
  };

  const checkSheet = (ws: WorksheetOptions, sheetName: string) => {
    (ws.rows ?? []).forEach((row, rowIdx) => {
      const rowHint = row.rowNumber ?? rowIdx + 1;
      for (const cell of row.cells ?? []) {
        const formula = typeof cell.formula === "string" ? cell.formula : cell.formula?.formula;
        if (!formula) continue;
        check(formula, `${sheetName}!${cell.reference ?? `row ${rowHint}`}`);
      }
    });
    for (const cf of ws.conditionalFormats ?? []) {
      for (const rule of cf.rules ?? []) {
        for (const formula of rule.formulas ?? []) {
          check(formula, `${sheetName} conditional format ${cf.sqref}`);
        }
      }
    }
    for (const dv of ws.dataValidations ?? []) {
      for (const formula of [dv.formula1, dv.formula2]) {
        if (formula) check(formula, `${sheetName} data validation ${dv.sqref}`);
      }
    }
  };

  options.worksheets?.forEach((ws, i) => checkSheet(ws, ws.name ?? `Sheet${i + 1}`));
  return issues;
}
