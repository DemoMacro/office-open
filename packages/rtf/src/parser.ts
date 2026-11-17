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
  const normalized = normalizeRtfSource(source);
  const tokens = tokenizeRtf(normalized);
  const sourceText = typeof normalized === "string" ? normalized : latin1Source(normalized);
  if (
    tokens[0]?.kind !== "group-start" ||
    tokens[1]?.kind !== "control" ||
    tokens[1].word !== "rtf"
  ) {
    throw new RtfParseError('RTF must begin with "{\\rtf"', tokens[0]?.position ?? 0, sourceText);
  }
  validateRootGroup(tokens, sourceText);
  return sanitizeOptionsText(parseRtfTokens(tokens, sourceText));
}

function normalizeRtfSource(source: string | Uint8Array): string | Uint8Array {
  if (typeof source === "string") return source.replace(/^\uFEFF/u, "");
  if (source.length >= 3 && source[0] === 0xef && source[1] === 0xbb && source[2] === 0xbf)
    return source.subarray(3);
  return source;
}

function validateRootGroup(tokens: ReturnType<typeof tokenizeRtf>, source: string): void {
  let depth = 0;
  let rootClosed = false;
  for (const token of tokens) {
    if (rootClosed) {
      if (token.kind !== "text" || token.value.trim() !== "")
        throw new RtfParseError("Content appears after the RTF root group", token.position, source);
      continue;
    }
    if (token.kind === "group-start") depth += 1;
    if (token.kind === "group-end") {
      depth -= 1;
      if (depth < 0)
        throw new RtfParseError("Unexpected closing RTF group", token.position, source);
      if (depth === 0) rootClosed = true;
    }
  }
  if (!rootClosed) throw new RtfParseError("RTF root group is not closed", source.length, source);
}
