import {
  applyPatchCommands,
  type OutputByType,
  type OutputType,
  type PatchCommand,
  type PatchOptions,
} from "@office-open/core";

import { generateDocument, parseDocument } from "./body";
import type { OdtDocumentOptions } from "./options";

/** One ordered canonical ODT edit. */
export type OdtPatchCommand = PatchCommand<OdtDocumentOptions>;

/** Options for parsing an ODT package, editing its canonical model, and regenerating it. */
export interface OdtPatchOptions<
  T extends OutputType = "uint8array",
> extends PatchOptions<OdtPatchCommand> {
  outputType?: T;
}

/**
 * Patch an existing ODT package through its canonical `OdtDocumentOptions`.
 *
 * Commands run in order and package overlay fields remain intact unless addressed.
 *
 * @publicApi
 */
export async function patchDocument<T extends OutputType = "uint8array">({
  commands,
  data,
  outputType,
  placeholderDelimiters,
}: OdtPatchOptions<T>): Promise<OutputByType[T]> {
  const options = await parseDocument(data);
  applyPatchCommands(options, commands, placeholderDelimiters);
  return generateDocument(options, (outputType ?? "uint8array") as T);
}
