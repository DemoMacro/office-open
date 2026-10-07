import type {
  ConditionalFormatRule,
  DefinedNameOptions,
  ExternalLinkOptions,
  FontOptions,
  RowOptions,
  WorksheetOptions,
} from "@office-open/xlsx";

import type { BiffVersion } from "../records";
export interface BoundSheet {
  readonly position: number;
  readonly state: WorksheetOptions["state"];
  readonly name: string;
}

export interface SheetState {
  rows: Map<number, RowOptions>;
  dimension?: string;
  extras: Partial<WorksheetOptions>;
}

export interface WorkbookState {
  version: BiffVersion;
  workbookLayout: "standard" | "biff4w";
  stream: Uint8Array;
  sharedStrings: string[];
  sheets: readonly BoundSheet[];
  fonts: FontOptions[];
  numberFormats: Map<number, string>;
  palette: string[];
  date1904: boolean;
  codepage: number;
  externalLinks: ExternalLinkOptions[];
  definedNames: DefinedNameOptions[];
  externSheetNames: Map<number, string>;
  styleTable: Map<number, { fontIndex?: number; numberFormatId?: number }>;
  drawingImages: { data: Uint8Array; type: "png" | "jpg" | "webp" | "wmf" | "emf" }[];
}

export interface FormulaLocation {
  row: number;
  column: number;
  style: number;
  expression?: string;
}

export interface TxoState {
  objectId: number;
  characterCount: number;
  runByteCount: number;
  textParts: TxoTextPart[];
  phase: "text" | "runs";
}

export interface TxoTextPart {
  readonly encoding: "utf-16le" | "windows-1252";
  readonly data: Uint8Array;
}

export interface TxoText {
  text: string;
}

export interface PendingConditional {
  reference: string;
  remaining: number;
  rules: ConditionalFormatRule[];
}
