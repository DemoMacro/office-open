import { applyPatchCommands } from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";

import { generateWorkbook } from "./generate";
import { parseWorkbook } from "./parse";
import type { WorkbookOptions } from "./parts/file";
import type { SheetDefinition } from "./parts/workbook";
import type { WorksheetOptions } from "./parts/worksheet";
import type { WorkbookPatchOptions } from "./patch-commands";

/**
 * Patch an existing XLSX package through its canonical `WorkbookOptions`.
 *
 * Commands run in order against the parsed model and unrelated canonical
 * fields — including round-trip-only package fields — remain intact.
 *
 * @publicApi
 */
export const patchWorkbook = async <T extends OutputType = "uint8array">({
  data,
  commands,
  outputType,
  packageVariant,
  placeholderDelimiters,
}: WorkbookPatchOptions<T>): Promise<OutputByType[T]> => {
  const options = await parseWorkbook(data);
  const originalWorksheets = [...(options.worksheets ?? [])];
  const originalSheetDefinitions = [...(options.sheetDefinitions ?? [])];
  applyPatchCommands(options, commands, placeholderDelimiters);
  reconcileSharedStrings(options);
  reconcileSheetDefinitions(options, originalWorksheets, originalSheetDefinitions);
  const output = await generateWorkbook(options, {
    type: outputType ?? "uint8array",
    packageVariant,
  });
  return output as OutputByType[T];
};

/** Keep shared-string payloads synchronized when a worksheet text command changes a shared cell. */
function reconcileSharedStrings(options: WorkbookOptions): void {
  if (!options.sharedStrings) return;
  for (const worksheet of options.worksheets ?? []) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        if (typeof cell.sharedIndex === "number" && typeof cell.value === "string") {
          options.sharedStrings[cell.sharedIndex] = cell.value;
        }
      }
    }
  }
}

/** Keep workbook sheet order and identity aligned after worksheet collection edits. */
function reconcileSheetDefinitions(
  options: WorkbookOptions,
  originalWorksheets: readonly WorksheetOptions[],
  originalDefinitions: readonly SheetDefinition[],
): void {
  const worksheets = options.worksheets ?? [];
  const paths = originalWorksheets.map(
    (worksheet, index) => worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`,
  );
  let nextSheetId = Math.max(
    0,
    ...originalDefinitions.map((definition) => definition.sheetId ?? 0),
  );
  let nextRelationship = Math.max(
    0,
    ...originalDefinitions.map((definition) => Number(definition.rId.replace(/\D/g, "") || 0)),
  );

  options.sheetDefinitions = worksheets.map((worksheet, worksheetIndex): SheetDefinition => {
    const sourcePath = worksheet.sourcePath ?? `xl/worksheets/sheet${worksheetIndex + 1}.xml`;
    const originalIndex = paths.indexOf(sourcePath);
    if (originalIndex >= 0) {
      const definition = originalDefinitions[originalIndex];
      if (definition) {
        return {
          ...definition,
          name: worksheet.name ?? definition.name,
          state: worksheet.state ?? definition.state,
        };
      }
    }
    nextSheetId += 1;
    nextRelationship += 1;
    return {
      name: worksheet.name ?? `Sheet${nextSheetId}`,
      sheetId: nextSheetId,
      rId: `rId${nextRelationship}`,
      state: worksheet.state,
    };
  });
}
