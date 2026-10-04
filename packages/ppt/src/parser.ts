import type { PresentationOptions } from "@office-open/pptx";

import { parsePresentation as parsePresentationData } from "./mappers/presentation";

export interface PptParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

export function parsePresentation(
  data: Uint8Array,
  options?: PptParseOptions,
): PresentationOptions {
  return parsePresentationData(data, options);
}
