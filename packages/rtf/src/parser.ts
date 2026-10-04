import type { DocumentOptions } from "@office-open/docx";

import { RtfParseError } from "./errors";
import { parseRtfTokens } from "./mappers/document";
import { tokenizeRtf } from "./tokenizer";

export type RtfParseOptions = Record<string, never>;

export function parseRtf(source: string, _options?: RtfParseOptions): DocumentOptions {
  const tokens = tokenizeRtf(source);
  if (
    tokens[0]?.kind !== "group-start" ||
    tokens[1]?.kind !== "control" ||
    tokens[1].word !== "rtf"
  ) {
    throw new RtfParseError('RTF must begin with "{\\rtf"', tokens[0]?.position ?? 0, source);
  }
  return parseRtfTokens(tokens, source);
}
