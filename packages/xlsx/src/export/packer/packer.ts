import type { WorkbookOptions } from "@file/file";
import { createPacker, OoxmlMimeType } from "@office-open/core";

import { compileWorkbook } from "../../compile/compiler";

export const Packer = createPacker<WorkbookOptions>({
  compile: (options, overrides, mediaLevel) => compileWorkbook(options, overrides, mediaLevel),
  mimeType: OoxmlMimeType.XLSX,
});
