import { Ajv, type ValidateFunction } from "ajv";
import addFormats from "ajv-formats";

import docxSchemaJson from "../../schemas/docx.schema.json";
import pptxSchemaJson from "../../schemas/pptx.schema.json";
import xlsxSchemaJson from "../../schemas/xlsx.schema.json";

/** Draft-07 JSON Schema describing a format's root options object. */
export type JsonSchema = { readonly [key: string]: unknown };

export const docxSchema = docxSchemaJson as unknown as JsonSchema;
export const pptxSchema = pptxSchemaJson as unknown as JsonSchema;
export const xlsxSchema = xlsxSchemaJson as unknown as JsonSchema;

export const SCHEMAS = {
  docx: docxSchema,
  pptx: pptxSchema,
  xlsx: xlsxSchema,
} as const;

export type DocumentType = "docx" | "pptx" | "xlsx";

const validators = new Map<DocumentType, ValidateFunction>();

function getValidator(type: DocumentType): ValidateFunction {
  let validate = validators.get(type);
  if (!validate) {
    const ajv = new Ajv({ allErrors: true, allowUnionTypes: true });
    addFormats(ajv);
    validate = ajv.compile(SCHEMAS[type]);
    validators.set(type, validate);
  }
  return validate;
}

export function validateDocumentInput(type: DocumentType, data: unknown): Record<string, unknown> {
  const validate = getValidator(type);
  if (!validate(data)) {
    const error = validate.errors?.[0];
    const path = (error?.instancePath ?? "").replaceAll("/", ".").replace(/^\./, "");
    throw new Error(
      `Invalid ${type} options${path ? ` at "${path}"` : ""}: ${error?.message ?? "invalid input"}`,
    );
  }
  return data as Record<string, unknown>;
}
