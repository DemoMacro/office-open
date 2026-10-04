import type { OfficeFormsOptions } from "@office-open/odf-schema";
import type { CellOptions, WorkbookOptions, WorksheetOptions } from "@office-open/xlsx";

export interface OdsCalculationSettings {
  caseSensitive?: boolean;
  automaticFindLabels?: boolean;
  regularExpressions?: boolean;
  wildcards?: boolean;
  nullDate?: string;
  iteration?: { enabled: boolean; steps?: number; maximumDifference?: number };
}

export interface OdsAnnotation {
  reference: string;
  paragraphs: string[];
  author?: string;
  date?: string;
}

export interface OdsCellGraphic {
  reference: string;
  href: string;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

export interface OdsSemanticsOptions {
  calculationSettings?: OdsCalculationSettings;
  forms?: OfficeFormsOptions;
}

export type OdsCellOptions = CellOptions & { covered?: boolean };
export interface OdsWorksheetOptions extends WorksheetOptions {
  annotations?: OdsAnnotation[];
  cellGraphics?: OdsCellGraphic[];
}
export interface OdsWorkbookOptions extends WorkbookOptions {
  odfSemantics?: OdsSemanticsOptions;
  worksheets?: OdsWorksheetOptions[];
}
