import type { DocumentOptions } from "@office-open/docx";

export function parseDocument(data: Uint8Array): DocumentOptions {
  void data;
  throw new Error("Legacy DOC parsing is not implemented yet");
}
