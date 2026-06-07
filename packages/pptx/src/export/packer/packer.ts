import type { File } from "@file/file";
import { createPacker, OoxmlMimeType } from "@office-open/core";

import { compilePresentation } from "../../compile/compiler";

export const Packer = createPacker<File>({
  compile: (file, overrides, mediaLevel) => compilePresentation(file, overrides, mediaLevel),
  mimeType: OoxmlMimeType.PPTX,
});
