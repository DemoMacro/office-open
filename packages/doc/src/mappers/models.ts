import type { RunOptions } from "@office-open/docx";

import type { NumberPair } from "../streams/pieces";

export interface Fib {
  nFib: number;
  ccpText: number;
  ccpFootnotes: number;
  ccpHeaders: number;
  ccpComments: number;
  ccpEndnotes: number;
  ccpTextboxes: number;
  ccpHeaderTextboxes: number;
  totalCharacters: number;
  clx: NumberPair;
  characterBinTable: NumberPair;
  paragraphBinTable: NumberPair;
  footnoteTable: NumberPair;
  headerTable: NumberPair;
  endnoteTable: NumberPair;
  textboxTable: NumberPair;
  headerTextboxTable: NumberPair;
  styleSheet: NumberPair;
  sectionTable: NumberPair;
  bookmarkNames: NumberPair;
  bookmarkStarts: NumberPair;
  bookmarkEnds: NumberPair;
  fields: NumberPair;
  comments: NumberPair;
  drawing: NumberPair;
  list: NumberPair;
  listOverrides: NumberPair;
}

export interface BinTableEntry {
  fcStart: number;
  fcEnd: number;
  page: number;
}

export interface CharacterRange {
  cpStart: number;
  cpEnd: number;
  properties: RunOptions;
}

export interface TextboxBoundary {
  start: number;
  end: number;
}

export interface ParagraphRange {
  cpStart: number;
  cpEnd: number;
  properties: LegacyParagraphProperties;
}

export interface LegacyParagraphProperties {
  alignment?: "left" | "center" | "right" | "both";
  indentLeft?: number;
  indentRight?: number;
  firstLine?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  spacingLine?: number;
  outlineLevel?: number;
  indent?: {
    left?: number;
    right?: number;
    firstLine?: number;
  };
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: "auto";
  };
  inTable?: boolean;
  rowEnd?: boolean;
  columnWidths?: number[];
  cellMerges?: Array<{ restart?: boolean; continue?: boolean }>;
}

export interface TextCharacter {
  value: string;
  cp: number;
  paragraphEnd?: boolean;
  hyperlink?: { url?: string; anchor?: string };
  revision?: { inserted: boolean };
}
