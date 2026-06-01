import type { File } from "@file/file";
import { createPacker, OoxmlMimeType } from "@office-open/core";

import { Compiler } from "./next-compiler";

const compiler = new Compiler();

/**
 * Exports documents to various output formats.
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * const buffer = await Packer.toBuffer(doc);
 * const blob = await Packer.toBlob(doc);
 * ```
 */
export const Packer = createPacker<File>({
  compile: (file, overrides, mediaLevel) => compiler.compile(file, overrides, mediaLevel),
  mimeType: OoxmlMimeType.DOCX,
});
