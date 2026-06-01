import type { File } from "@file/file";
/**
 * Packer module — export API for XLSX files.
 *
 * @module
 */
import { createPacker, OoxmlMimeType } from "@office-open/core";

import { Compiler } from "./next-compiler";

const compiler = new Compiler();

export const Packer = createPacker<File>({
  compile: (file, overrides, mediaLevel) => compiler.compile(file, overrides, mediaLevel),
  mimeType: OoxmlMimeType.XLSX,
});
