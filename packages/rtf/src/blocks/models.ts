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
  "bold" | "italic" | "underline" | "color" | "size" | "font"
>;

export type ParagraphDraft = {
  options: ParagraphOptions;
  runs: RunOptions[];
  format: RunFormat;
  segments: (string | NonNullable<RunOptions["children"]>[number])[];
};

export type CellDraft = { blocks: SectionChild[] };

export type TableDraft = {
  target: TableOptions;
  rows: TableRowOptions[];
  cells: TableCellOptions[];
  cell: CellDraft;
  cellOpen: boolean;
  cellEnded: boolean;
  parentBlocks: SectionChild[];
};
