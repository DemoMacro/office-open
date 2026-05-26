import type { File } from "@file/file";
import { createPacker, OoxmlMimeType } from "@office-open/core";
export { PrettifyType } from "@office-open/core";

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
 * const buffer = await Packer.toBuffer(doc, PrettifyType.WITH_2_BLANKS);
 * ```
 */
export const Packer = createPacker<File>({
  compile: (file, prettify, overrides) => compiler.compile(file, prettify, overrides),
  mimeType: OoxmlMimeType.DOCX,
});
