import { readFile } from "node:fs/promises";

import { detectOfficeFormat, type OfficeFormatInfo } from "./formats";

export async function detectOfficeFile(path: string): Promise<OfficeFormatInfo> {
  return detectOfficeFormat(new Uint8Array(await readFile(path)));
}
