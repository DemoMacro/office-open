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
  reconcileInsertedWorksheets(options, originalWorksheets);
  reconcileSheetDefinitions(options, originalWorksheets, originalSheetDefinitions);
  removeDeletedWorksheetContentTypes(options, originalWorksheets);
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

/** Drop content-type declarations for worksheets removed by a patch command. */
function removeDeletedWorksheetContentTypes(
  options: WorkbookOptions,
  originalWorksheets: readonly WorksheetOptions[],
): void {
  const retainedPaths = new Set(
    (options.worksheets ?? []).map((worksheet) => worksheet.sourcePath?.toLowerCase()),
  );
  options.contentTypes = options.contentTypes && {
    ...options.contentTypes,
    overrides: options.contentTypes.overrides.filter(
      (override) =>
        !matchesOriginalWorksheet(override.partName, originalWorksheets) ||
        retainedPaths.has(override.partName.toLowerCase()),
    ),
  };
}

/** Assign package paths and workbook identities to worksheets inserted by patch commands. */
function reconcileInsertedWorksheets(
  options: WorkbookOptions,
  originalWorksheets: readonly WorksheetOptions[],
): void {
  const usedPaths = new Set(
    originalWorksheets.map((worksheet, index) =>
      (worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`).toLowerCase(),
    ),
  );
  const usedSheetIds = new Set<number>();
  const sheetFamilies = [
    [...originalWorksheets, ...(options.worksheets ?? [])],
    options.chartsheets ?? [],
    options.dialogsheets ?? [],
    options.macrosheets ?? [],
  ];
  for (const sheets of sheetFamilies) {
    for (const sheet of sheets) {
      if (sheet.sheetId !== undefined) usedSheetIds.add(sheet.sheetId);
    }
  }
  let nextSheetId = Math.max(0, ...usedSheetIds) + 1;

  for (const worksheet of options.worksheets ?? []) {
    if (worksheet.sourcePath === undefined) {
      let part = 1;
      while (usedPaths.has(`xl/worksheets/sheet${part}.xml`)) part += 1;
      worksheet.sourcePath = `xl/worksheets/sheet${part}.xml`;
      usedPaths.add(worksheet.sourcePath.toLowerCase());
    }
    if (worksheet.sheetId === undefined) {
      while (usedSheetIds.has(nextSheetId)) nextSheetId += 1;
      worksheet.sheetId = nextSheetId;
      usedSheetIds.add(nextSheetId);
      nextSheetId += 1;
    }
  }
}

/** Match a leading-slash OPC part name to a worksheet's package-relative source path. */
function matchesOriginalWorksheet(
  partName: string,
  originalWorksheets: readonly WorksheetOptions[],
): boolean {
  const normalizedPartName = partName.replace(/^\/+/, "").toLowerCase();
  return originalWorksheets.some((worksheet, index) => {
    const sourcePath = worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`;
    return sourcePath.toLowerCase() === normalizedPartName;
  });
}

/** Keep workbook sheet order and identity aligned after worksheet collection edits. */
function reconcileSheetDefinitions(
  options: WorkbookOptions,
  originalWorksheets: readonly WorksheetOptions[],
  originalDefinitions: readonly SheetDefinition[],
): void {
  const worksheets = options.worksheets ?? [];
  const usedSheetIds = new Set<number>([
    ...originalDefinitions.map((definition) => definition.sheetId ?? 0),
    ...(options.chartsheets ?? []).flatMap((sheet) => (sheet.sheetId ? [sheet.sheetId] : [])),
    ...(options.dialogsheets ?? []).flatMap((sheet) => (sheet.sheetId ? [sheet.sheetId] : [])),
    ...(options.macrosheets ?? []).flatMap((sheet) => (sheet.sheetId ? [sheet.sheetId] : [])),
  ]);
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
    const sheetId =
      worksheet.sheetId ??
      (() => {
        while (usedSheetIds.has(nextSheetId)) nextSheetId += 1;
        return nextSheetId;
      })();
    usedSheetIds.add(sheetId);
    nextRelationship += 1;
    return {
      name: worksheet.name ?? `Sheet${nextSheetId}`,
      sheetId,
      rId: `rId${nextRelationship}`,
      state: worksheet.state,
    };
  });
}
