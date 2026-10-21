import { applyPatchCommands } from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";

import { generateDocument } from "../generate";
import { parseDocument } from "../parse";
import type { DocumentPatchOptions } from "./commands";

/**
 * Patch an existing DOCX package through its canonical `DocumentOptions`.
 *
 * Commands run in order against the parsed model and unrelated canonical
 * fields — including round-trip-only package fields — remain intact.
 *
 * @publicApi
 */
export const patchDocument = async <T extends OutputType = "uint8array">({
  data,
  commands,
  outputType,
  packageVariant,
  placeholderDelimiters,
}: DocumentPatchOptions<T>): Promise<OutputByType[T]> => {
  const options = await parseDocument(data);
  applyPatchCommands(options, commands, placeholderDelimiters);
  const output = await generateDocument(options, {
    type: outputType ?? "uint8array",
    packageVariant,
  });
  return output as OutputByType[T];
};
