/**
 * The committed draft-07 schemas and their ajv validators.
 *
 * @module
 */
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

/** Max validation issues surfaced in one error message. */
const MAX_REPORTED_ERRORS = 5;

/**
 * Validate document options against the full format schema.
 * Aggregates up to five `instancePath`-qualified issues so callers (and LLMs)
 * can fix several problems per iteration instead of one.
 */
export function validateDocumentInput(type: DocumentType, data: unknown): Record<string, unknown> {
  const validate = getValidator(type);
  if (!validate(data)) {
    const errors = validate.errors ?? [];
    const lines = errors.slice(0, MAX_REPORTED_ERRORS).map((error) => {
      const path = (error.instancePath ?? "").replaceAll("/", ".").replace(/^\./, "");
      const message = error.message ?? "invalid input";
      return path ? `at "${path}": ${message}` : message;
    });
    const more =
      errors.length > MAX_REPORTED_ERRORS ? ` (+${errors.length - MAX_REPORTED_ERRORS} more)` : "";
    throw new Error(
      `Invalid ${type} options: ${lines.length > 0 ? lines.join("; ") : "invalid input"}${more}`,
    );
  }
  return data as Record<string, unknown>;
}
