export type FormulaFunctionCategory =
  | "cube"
  | "database"
  | "dateTime"
  | "dynamicArray"
  | "engineering"
  | "financial"
  | "information"
  | "lambda"
  | "logical"
  | "lookupReference"
  | "math"
  | "openFormula"
  | "other"
  | "statistical"
  | "text";

export type FormulaFunctionExternalContext =
  | "cube"
  | "dde"
  | "detectLanguage"
  | "externalWorkbook"
  | "image"
  | "linkedData"
  | "pivot"
  | "stockHistory"
  | "web"
  | "translation"
  | "macro"
  | "rtd";

export interface FormulaFunctionCatalogEntry {
  name: string;
  category: FormulaFunctionCategory;
  minArgs: number;
  maxArgs: number;
  astOnly?: boolean;
  volatile?: boolean;
  dynamicDependencies?: boolean;
  acceptsErrors?: boolean;
  externalContext?: FormulaFunctionExternalContext;
}
