import { sanitizeOptionsText } from "@office-open/core";
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
  return sanitizeOptionsText(parseRtfTokens(tokens, sourceText));
}
