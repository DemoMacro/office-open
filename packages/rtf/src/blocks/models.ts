import type {
  ParagraphOptions,
  RunOptions,
  SectionChild,
  TableCellOptions,
  TableOptions,
  TableRowOptions,
} from "@office-open/docx";

export type RunFormat = Pick<
  RunOptions,
  | "additionRsid"
  | "bold"
  | "italic"
  | "underline"
  | "color"
  | "size"
  | "font"
  | "strike"
  | "verticalAlign"
  | "language"
  | "characterSpacing"
  | "kern"
  | "allCaps"
  | "smallCaps"
  | "outline"
  | "shadow"
  | "vanish"
  | "emboss"
  | "scale"
  | "style"
  | "highlight"
>;

export type ParagraphDraft = {
  options: ParagraphOptions;
  runs: RunOptions[];
  format: RunFormat;
  segments: (string | NonNullable<RunOptions["children"]>[number])[];
};

export type CellDraft = {
  blocks: SectionChild[];
  properties: Partial<TableCellOptions>;
};

export type TableDraft = {
  target: TableOptions;
  rows: TableRowOptions[];
  cells: TableCellOptions[];
  cell: CellDraft;
  rowProperties: Partial<TableRowOptions>;
  boundaries: number[];
  cellDefinitions: Partial<TableCellOptions>[];
  cellOpen: boolean;
  cellEnded: boolean;
  parentBlocks: SectionChild[];
};
