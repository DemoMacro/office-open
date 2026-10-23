import {
  applyPatchCommands,
  convertOutput,
  toUint8ArrayAsync,
  type OutputByType,
  type OutputType,
  type PatchCommand,
  type PatchOptions,
} from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";

import { generateDocument } from "./generate";
import { parseDocument } from "./parser";

/** One ordered canonical RTF edit. */
export type RtfPatchCommand = PatchCommand<DocumentOptions>;

/** Options for parsing an RTF file, editing its canonical model, and regenerating it. */
export interface RtfPatchOptions<
  T extends OutputType = "uint8array",
> extends PatchOptions<RtfPatchCommand> {
  outputType?: T;
}

/**
 * Patch an existing RTF file through its canonical `DocumentOptions`.
 *
 * The operation is a canonical edit and regeneration, not source-preserving
 * rewriting; fields the RTF writer cannot express retain `RtfGenerateError`.
 *
 * @publicApi
 */
export async function patchDocument<T extends OutputType = "uint8array">({
  commands,
  data,
  outputType,
  placeholderDelimiters,
}: RtfPatchOptions<T>): Promise<OutputByType[T]> {
  const options = parseDocument(await toUint8ArrayAsync(data));
  applyPatchCommands(options, commands, placeholderDelimiters);
  return convertRtfOutput(generateDocument(options), outputType);
}

function convertRtfOutput<T extends OutputType>(data: Uint8Array, outputType?: T): OutputByType[T] {
  return convertOutput(data, (outputType ?? "uint8array") as T, "application/rtf");
}
