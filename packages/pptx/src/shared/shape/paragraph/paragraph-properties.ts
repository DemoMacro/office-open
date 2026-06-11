/**
 * Paragraph properties types for PPTX text paragraphs.
 *
 * @module
 */

export type TextAlignment = "left" | "center" | "right" | "justify";

export type BulletCharOptions = {
  type: "char";
  char?: string;
  color?: string;
  size?: number;
};

export type BulletAutoNumOptions = {
  type: "autoNum";
  format?: string;
  startAt?: number;
  color?: string;
  size?: number;
};

export type BulletNoneOption = {
  type: "none";
};

export type BulletOptions = BulletCharOptions | BulletAutoNumOptions | BulletNoneOption;

export interface ParagraphPropertiesOptions {
  alignment?: TextAlignment;
  indentLevel?: number;
  marginBottom?: number;
  marginTop?: number;
  bullet?: BulletOptions;
  lineSpacing?: number;
  lineSpacingPoints?: number;
  marginIndent?: number;
  marginRight?: number;
  defTabSize?: number;
  fontAlignment?: "auto" | "t" | "ctr" | "b" | "base";
}
