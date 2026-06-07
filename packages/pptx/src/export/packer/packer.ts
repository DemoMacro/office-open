import type { PresentationOptions } from "@file/file";
import { createPacker, OoxmlMimeType } from "@office-open/core";

import { compilePresentation } from "../../compile/compiler";

export const Packer = createPacker<PresentationOptions>({
  compile: (options, overrides, mediaLevel) => compilePresentation(options, overrides, mediaLevel),
  mimeType: OoxmlMimeType.PPTX,
});
