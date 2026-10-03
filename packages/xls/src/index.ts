import type { WorkbookOptions } from "@office-open/xlsx";

export function parseWorkbook(data: Uint8Array): WorkbookOptions {
  void data;
  throw new Error("Legacy XLS parsing is not implemented yet");
}
