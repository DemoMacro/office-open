import type { DocumentOptions } from "@office-open/docx";
import type { OfficeFormsOptions } from "@office-open/odf-schema";

export interface OdtSequenceDeclaration {
  name: string;
  displayOutlineLevel: number;
  separationCharacter?: string;
}

export interface OdtVariableDeclaration {
  name: string;
  valueType: "float" | "percentage" | "currency" | "date" | "time" | "boolean" | "string";
}

export interface OdtTrackedChange {
  id: string;
  kind: "insertion" | "deletion" | "format-change";
  author?: string;
  date?: string;
}

export interface OdtSemanticsOptions {
  forms?: OfficeFormsOptions;
  sequenceDeclarations?: OdtSequenceDeclaration[];
  variableDeclarations?: OdtVariableDeclaration[];
  trackedChanges?: { trackChanges?: boolean; changes: OdtTrackedChange[] };
}

export interface OdtDocumentOptions extends DocumentOptions {
  odfSemantics?: OdtSemanticsOptions;
}
