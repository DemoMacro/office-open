import type { PresentationOptions } from "@office-open/pptx";

export function parsePresentation(data: Uint8Array): PresentationOptions {
  void data;
  throw new Error("Legacy PPT parsing is not implemented yet");
}
