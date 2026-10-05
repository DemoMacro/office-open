import type { DocumentOptions } from "@office-open/docx";

import { writeRtf } from "./writer/document";

export function generateRtf(options: DocumentOptions): Uint8Array {
  return new TextEncoder().encode(writeRtf(options));
}
