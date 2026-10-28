import type { FormulaAst, FormulaDialect, FormulaWorkbookIdentity } from "./ast";
import type { FormulaMatrix, FormulaScalar, FormulaValue } from "./value";

export interface FormulaCellSnapshot {
  value?: FormulaScalar;
  error?: true;
  formula?: string;
}

export interface FormulaCellEntry extends FormulaCellSnapshot {
  row: number;
  column: number;
}

export interface FormulaLambdaDefinition {
  parameterNames: string[];
  body: FormulaAst;
  variables: ReadonlyMap<string, FormulaValue>;
}

export interface FormulaCallerContext {
  readonly sheet?: number;
  readonly row?: number;
  readonly column?: number;
}

export interface FormulaHostFunctionRequest {
  readonly name: string;
  readonly values: readonly FormulaValue[];
  readonly nodes?: readonly FormulaAst[];
  readonly caller: FormulaCallerContext;
}

export interface PivotFieldSelector {
  field: string;
  item: string;
}

export interface GetPivotDataRequest {
  dataField: string;
  pivotTable: string;
  selectors: readonly PivotFieldSelector[];
}

export interface RealTimeDataRequest {
  programId: string;
  server: string;
  topics: readonly string[];
}

export interface ImageFunctionRequest {
  source: string;
  altText?: string;
  sizing?: number;
  height?: number;
  width?: number;
}

export interface StockHistoryFunctionRequest {
  symbol: string;
  startDate: FormulaScalar;
  endDate?: FormulaScalar;
  interval?: number;
  headers?: boolean;
  properties?: readonly number[];
  identifier?: string;
}

export interface TranslateTextFunctionRequest {
  text: FormulaScalar;
  sourceLanguage?: string;
  targetLanguage?: string;
}

export interface DetectLanguageFunctionRequest {
  text: readonly FormulaScalar[];
}

export interface FieldValueFunctionRequest {
  value: FormulaScalar;
  field: string;
}

export interface MacroCallRequest {
  library: string;
  procedure: string;
  typeText?: string;
  arguments: readonly FormulaScalar[];
}

export interface MacroRegistrationRequest {
  library: string;
  procedure: string;
  typeText?: string;
}

export interface SqlRequest {
  connectionString: string;
  query?: string;
  outputReference?: string;
  driverPrompt?: number;
  columnNames?: boolean;
  rowNames?: boolean;
  maxRows?: number;
  maxChanges?: number;
  timeout?: number;
}

export interface CubeMemberRequest {
  connection: string;
  expression: string;
  caption?: string;
}

export interface CubeKpiRequest extends CubeMemberRequest {
  property: string;
}

export interface CubeRankedRequest {
  connection: string;
  setExpression: string;
  rank: number;
  caption?: string;
}

export interface CubeSetRequest {
  connection: string;
  setExpression: string;
  caption?: string;
  sortOrder?: number;
  sortBy?: string;
}

export interface CubeValueRequest {
  connection: string;
  expressions: readonly string[];
}

export interface DdeRequest {
  server: string;
  file: string;
  item: string;
  mode?: string;
}

export interface XmlXPathRequest {
  xml: string;
  expression: string;
}

export interface FormulaSharedGroup {
  readonly masterRow: number;
  readonly masterColumn: number;
  readonly formula: string;
}

export interface FormulaWorkbookContext {
  resolveDynamic?(request: FormulaHostFunctionRequest): FormulaValue | undefined;
  resolveLookup?(request: FormulaHostFunctionRequest): FormulaValue | undefined;
  getPivotValue?(request: GetPivotDataRequest): FormulaValue | undefined;
  getRealTimeData?(request: RealTimeDataRequest): FormulaValue | undefined;
  getImage?(request: ImageFunctionRequest): FormulaValue | undefined;
  getStockHistory?(request: StockHistoryFunctionRequest): FormulaValue | undefined;
  translateText?(request: TranslateTextFunctionRequest): FormulaValue | undefined;
  detectLanguage?(request: DetectLanguageFunctionRequest): FormulaValue | undefined;
  getFieldValue?(request: FieldValueFunctionRequest): FormulaValue | undefined;
  fetchWebService?(url: string): FormulaValue | undefined;
  callMacro?(request: MacroCallRequest): FormulaValue | undefined;
  registerMacro?(request: MacroRegistrationRequest): FormulaValue | undefined;
  executeSql?(request: SqlRequest): FormulaValue | undefined;
  getCubeMember?(request: CubeMemberRequest): FormulaValue | undefined;
  getCubeKpiMember?(request: CubeKpiRequest): FormulaValue | undefined;
  getCubeMemberProperty?(request: CubeKpiRequest): FormulaValue | undefined;
  getCubeRankedMember?(request: CubeRankedRequest): FormulaValue | undefined;
  getCubeSet?(request: CubeSetRequest): FormulaValue | undefined;
  getCubeSetCount?(request: { set: FormulaValue }): FormulaValue | undefined;
  getCubeValue?(request: CubeValueRequest): FormulaValue | undefined;
  getDdeData?(request: DdeRequest): FormulaValue | undefined;
  evaluateXmlXPath?(request: XmlXPathRequest): FormulaValue | undefined;
  getWorkbook?(workbook: FormulaWorkbookIdentity): FormulaWorkbookContext | undefined;
  getSheetId(name: string): number | undefined;
  getSheetName(id: number): string | undefined;
  getCell(sheet: number, row: number, column: number): FormulaCellSnapshot | undefined;
  getCells?(
    sheet: number,
    bounds: {
      firstRow: number;
      firstColumn: number;
      lastRow: number;
      lastColumn: number;
    },
  ): Iterable<FormulaCellEntry>;
  getDefinedFormula(name: string, sheetId?: number): string | undefined;
  getTableName(name: string):
    | {
        id: number;
        sheetId: number;
        range: { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number };
        columnNames?: string[];
        hasHeaders?: boolean;
        hasTotals?: boolean;
      }
    | undefined;
  getRowCount(sheet: number): number;
  getColumnCount(sheet: number): number;
  getSpillRange?(
    sheet: number,
    row: number,
    column: number,
  ): { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number } | undefined;
  getEuroExchangeRate?(currency: string): number | undefined;
  isRowHidden?(sheet: number, row: number): boolean;
  now?: () => Date;
  random?: () => number;
  locale?: string;
  dateSystem?: "1900" | "1904";
  textEncoding?: FormulaTextEncoding;
}

export interface FormulaEvaluationOptions {
  dialect?: FormulaDialect;
  dateSystem?: "1900" | "1904";
  errors?: "object" | "serialized";
  sheet?: number;
  row?: number;
  column?: number;
}

export interface FormulaTextEncoding {
  codePage?: number;
  encode?(text: string): Uint8Array;
  decode?(bytes: Uint8Array): string;
}

export interface FormulaEvaluationContext extends FormulaEvaluationOptions {
  readonly workbook: FormulaWorkbookContext;
  readonly variables?: ReadonlyMap<string, FormulaValue>;
  readonly lambdaScope: Map<string, FormulaLambdaDefinition>;
  invokeLambdaValue?(handle: string, args: readonly FormulaValue[]): FormulaValue;
}

export function createFormulaContext(
  workbook: FormulaWorkbookContext,
  options: FormulaEvaluationOptions = {},
): FormulaEvaluationContext {
  return { ...options, workbook, lambdaScope: new Map() };
}

export type { FormulaMatrix, FormulaScalar };
