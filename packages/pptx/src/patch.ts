import { applyPatchCommands } from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";

import { generatePresentation } from "./generate";
import { parsePresentation } from "./parse";
import type { PresentationPatchOptions } from "./patch-commands";

/**
 * Patch an existing PPTX package through its canonical `PresentationOptions`.
 *
 * Commands run in order against the parsed model and unrelated canonical
 * fields — including round-trip-only package fields — remain intact.
 *
 * @publicApi
 */
export const patchPresentation = async <T extends OutputType = "uint8array">({
  data,
  commands,
  outputType,
  packageVariant,
  placeholderDelimiters,
}: PresentationPatchOptions<T>): Promise<OutputByType[T]> => {
  const options = await parsePresentation(data);
  applyPatchCommands(options, commands, placeholderDelimiters);
  const output = await generatePresentation(options, {
    type: outputType ?? "uint8array",
    packageVariant,
  });
  return output as OutputByType[T];
};
