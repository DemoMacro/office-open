import type {
  DocumentOptions,
  ParagraphChild,
  ParagraphOptions,
  SectionChild,
  SectionOptions,
} from "@office-open/docx";
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

export interface OdtFieldValue {
  name: string;
  value?: string;
}

export interface OdtVariableSet {
  name: string;
  valueType: OdtVariableDeclaration["valueType"];
  value?: string;
  display?: string;
  formula?: string;
}

export type OdtInlineChild =
  | { annotation: OdtAnnotation }
  | { bookmarkEnd: { name: string } }
  | { bookmarkReference: OdtBookmarkReference }
  | { bibliographyMark: OdtBibliographyMark }
  | { chapter: OdtChapter }
  | { expression: OdtExpression }
  | { hiddenText: OdtHiddenText }
  | { referenceMarkEnd: { name: string } }
  | { referenceMarkStart: { name: string } }
  | { referenceReference: OdtReferenceReference }
  | { sequence: OdtSequence }
  | { shape: OdtShape }
  | { shapeHyperlink: OdtShapeHyperlink }
  | { variableSet: OdtVariableSet };

export interface OdtParagraphOptions extends Omit<ParagraphOptions, "children"> {
  children?: (string | ParagraphChild | OdtInlineChild)[];
}

export type OdtSectionChild =
  | Exclude<SectionChild, { paragraph: unknown }>
  | { paragraph: string | OdtParagraphOptions }
  | { bibliography: OdtBibliography }
  | { index: OdtIndex }
  | { shape: OdtShape }
  | { shapeHyperlink: OdtShapeHyperlink };

export interface OdtSectionOptions extends Omit<SectionOptions, "children"> {
  children: OdtSectionChild[];
}

export interface OdtBookmarkReference {
  name: string;
  referenceFormat?: string;
}

export interface OdtSequence {
  name: string;
  referenceName?: string;
  display?: string;
  formula?: string;
}

export interface OdtChapter {
  display?: string;
  outlineLevel?: number;
}

export interface OdtExpression {
  formula?: string;
  valueType?: OdtVariableDeclaration["valueType"];
  value?: string;
  display?: string;
}

export interface OdtHiddenText {
  condition: string;
  content: string;
  hidden?: boolean;
  fixed?: boolean;
}

export interface OdtReferenceMarkEnd {
  name: string;
}

export interface OdtBibliographyMark {
  type: string;
  identifier?: string;
  fields?: OdtFieldValue[];
}

export interface OdtAnnotation {
  paragraphs: string[];
  author?: string;
  date?: string;
  display?: boolean;
}

export interface OdtIndexSource {
  scope?: string;
  useCaption?: boolean;
  relativeTabStopPosition?: boolean;
  title?: string;
  titleStyleName?: string;
}

export interface OdtIndex {
  name?: string;
  styleName?: string;
  source: OdtIndexSource;
  paragraphs: string[];
}

export interface OdtBibliography {
  name?: string;
  styleName?: string;
  source: OdtIndexSource;
  paragraphs: string[];
}

export interface OdtShape {
  geometry: "ellipse" | "rect";
  name?: string;
  styleName?: string;
  x?: string;
  y?: string;
  width?: string;
  height?: string;
}

export interface OdtReferenceReference {
  name: string;
  referenceFormat?: string;
}

export interface OdtShapeHyperlink {
  href: string;
  shapes: OdtShape[];
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
