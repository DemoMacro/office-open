import { sanitizeXmlText } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";

import { RtfParseError } from "./errors";
import { parseRtfTokens } from "./mappers/document";
import { latin1Source, tokenizeRtf } from "./tokenizer";

export type RtfParseOptions = Record<string, never>;

export function parseDocument(
  source: string | Uint8Array,
  _options?: RtfParseOptions,
): DocumentOptions {
  const tokens = tokenizeRtf(source);
  const sourceText = typeof source === "string" ? source : latin1Source(source);
  if (
    tokens[0]?.kind !== "group-start" ||
    tokens[1]?.kind !== "control" ||
    tokens[1].word !== "rtf"
  ) {
    throw new RtfParseError('RTF must begin with "{\\rtf"', tokens[0]?.position ?? 0, sourceText);
  }
  return sanitizeDocumentText(parseRtfTokens(tokens, sourceText));
}

function sanitizeDocumentText<T>(value: T): T {
  if (typeof value === "string") return sanitizeXmlText(value) as T;
  if (value instanceof Uint8Array || value instanceof Date) return value;
  if (Array.isArray(value)) return value.map(sanitizeDocumentText) as T;
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, child]) => [key, sanitizeDocumentText(child)]),
    ) as T;
  }
  return value;
}
