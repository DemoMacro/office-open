import type { DocumentOptions } from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";

export function parseOdt(data: Uint8Array): DocumentOptions {
  void data;
  throw new Error("ODT parsing is not implemented yet");
}

export function parseOds(data: Uint8Array): WorkbookOptions {
  void data;
  throw new Error("ODS parsing is not implemented yet");
}

export function parseOdp(data: Uint8Array): PresentationOptions {
  void data;
  throw new Error("ODP parsing is not implemented yet");
}

export function generateOdt(options: DocumentOptions): Uint8Array {
  void options;
  throw new Error("ODT generation is not implemented yet");
}

export function generateOds(options: WorkbookOptions): Uint8Array {
  void options;
  throw new Error("ODS generation is not implemented yet");
}

export function generateOdp(options: PresentationOptions): Uint8Array {
  void options;
  throw new Error("ODP generation is not implemented yet");
}
