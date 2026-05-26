import type { File } from "@file/file";
import { createPacker, OoxmlMimeType } from "@office-open/core";
export { PrettifyType } from "@office-open/core";

import { Compiler } from "./next-compiler";

const compiler = new Compiler();

export const Packer = createPacker<File>({
  compile: (file, prettify, overrides) => compiler.compile(file, prettify, overrides),
  mimeType: OoxmlMimeType.PPTX,
});
