/**
 * Namespace configuration for XML patch operations.
 *
 * Parameterises element names so the same patch algorithm works for both
 * DOCX (`w:*`) and PPTX (`a:*`) documents.
 */

export interface XmlNamespaceConfig {
  paragraph: string;
  run: string;
  text: string;
  runProperties: string;
}

export const DOCX_NS: XmlNamespaceConfig = {
  paragraph: "w:p",
  run: "w:r",
  text: "w:t",
  runProperties: "w:rPr",
};

export const PPTX_NS: XmlNamespaceConfig = {
  paragraph: "a:p",
  run: "a:r",
  text: "a:t",
  runProperties: "a:rPr",
};
