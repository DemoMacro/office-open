import { docxSchema } from "./docx";
import { pptxSchema } from "./pptx";
import { xlsxSchema } from "./xlsx";

export { docxSchema } from "./docx";
export { pptxSchema } from "./pptx";
export { xlsxSchema } from "./xlsx";

export const SCHEMAS = {
  docx: docxSchema,
  pptx: pptxSchema,
  xlsx: xlsxSchema,
} as const;

export type DocumentType = "docx" | "pptx" | "xlsx";

export function validateDocumentInput(type: DocumentType, data: unknown): Record<string, unknown> {
  const result = SCHEMAS[type].safeParse(data);
  if (!result.success) {
    const issue = result.error.issues[0];
    const path = issue.path.join(".");
    throw new Error(`Invalid ${type} options${path ? ` at "${path}"` : ""}: ${issue.message}`);
  }
  return result.data as Record<string, unknown>;
}
