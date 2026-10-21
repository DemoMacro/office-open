import {
  applyPatchCommands,
  type OutputByType,
  type OutputType,
  type PatchCommand,
  type PatchOptions,
} from "@office-open/core";

import type { OdsWorkbookOptions } from "./options";
import { generateWorkbook, parseWorkbook } from "./workbook";

/** One ordered canonical ODS edit. */
export type OdsPatchCommand = PatchCommand<OdsWorkbookOptions>;

/** Options for parsing an ODS package, editing its canonical model, and regenerating it. */
export interface OdsPatchOptions<
  T extends OutputType = "uint8array",
> extends PatchOptions<OdsPatchCommand> {
  outputType?: T;
}

/**
 * Patch an existing ODS package through its canonical `OdsWorkbookOptions`.
 *
 * Commands run in order and package overlay fields remain intact unless addressed.
 *
 * @publicApi
 */
export async function patchWorkbook<T extends OutputType = "uint8array">({
  commands,
  data,
  outputType,
  placeholderDelimiters,
}: OdsPatchOptions<T>): Promise<OutputByType[T]> {
  const options = await parseWorkbook(data);
  applyPatchCommands(options, commands, placeholderDelimiters);
  return generateWorkbook(options, (outputType ?? "uint8array") as T);
}
