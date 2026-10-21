import {
  applyPatchCommands,
  type OutputByType,
  type OutputType,
  type PatchCommand,
  type PatchOptions,
} from "@office-open/core";

import type { OdpPresentationOptions } from "./options";
import { generatePresentation, parsePresentation } from "./presentation";

/** One ordered canonical ODP edit. */
export type OdpPatchCommand = PatchCommand<OdpPresentationOptions>;

/** Options for parsing an ODP package, editing its canonical model, and regenerating it. */
export interface OdpPatchOptions<
  T extends OutputType = "uint8array",
> extends PatchOptions<OdpPatchCommand> {
  outputType?: T;
}

/**
 * Patch an existing ODP package through its canonical `OdpPresentationOptions`.
 *
 * Commands run in order and package overlay fields remain intact unless addressed.
 *
 * @publicApi
 */
export async function patchPresentation<T extends OutputType = "uint8array">({
  commands,
  data,
  outputType,
  placeholderDelimiters,
}: OdpPatchOptions<T>): Promise<OutputByType[T]> {
  const options = await parsePresentation(data);
  applyPatchCommands(options, commands, placeholderDelimiters);
  return generatePresentation(options, (outputType ?? "uint8array") as T);
}
