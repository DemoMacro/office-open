import { applyPatchCommands } from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";

import { generateWorkbook } from "./generate";
import { parseWorkbook } from "./parse";
import type { WorkbookOptions } from "./parts/file";
import type { SheetDefinition } from "./parts/workbook";
import type { CellOptions, WorksheetOptions } from "./parts/worksheet";
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
  const sharedStrings = options.sharedStrings;
  const cells: Array<{ cell: CellOptions; entry: SharedStringEntry }> = [];
  let referenceCount = 0;
  let changed = false;
  for (const worksheet of options.worksheets ?? []) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        if (usesSharedStrings(cell)) {
          referenceCount += 1;
          const entry = sharedStringEntry(cell);
          if (entry === undefined) continue;
          const index = cell.sharedIndex;
          const current =
            typeof index === "number" && index < sharedStrings.length
              ? sharedStrings[index]
              : undefined;
          if (sharedStringKey(current) !== sharedStringKey(entry)) changed = true;
          cells.push({ cell, entry });
        }
      }
    }
  }

  if (changed) {
    const indices = new Map<string, number>();
    const reconciled: SharedStringEntry[] = [];
    for (const { cell, entry } of cells) {
      const key = sharedStringKey(entry)!;
      let index = indices.get(key);
      if (index === undefined) {
        index = reconciled.length;
        indices.set(key, index);
        reconciled.push(entry);
      }
      cell.sharedIndex = index;
    }
    options.sharedStrings = reconciled;
  }
  if (options.sharedStringsCount !== undefined) options.sharedStringsCount = referenceCount;
  if (options.sharedStringsUniqueCount !== undefined)
    options.sharedStringsUniqueCount = options.sharedStrings.length;
}

type SharedStringEntry = NonNullable<WorkbookOptions["sharedStrings"]>[number];

function sharedStringEntry(cell: CellOptions): SharedStringEntry | undefined {
  if (typeof cell.value === "string") return cell.value;
  if (cell.value === null || typeof cell.value !== "object" || cell.value instanceof Date)
    return undefined;
  return cell.value;
}

function sharedStringKey(entry: SharedStringEntry | undefined): string {
  return JSON.stringify(entry);
}

function usesSharedStrings(cell: CellOptions): boolean {
  if (cell.inline || cell.error !== undefined) return false;
  if (typeof cell.value === "string") {
    return cell.typeRaw !== "n" && cell.typeRaw !== "str" && cell.typeRaw !== "d";
  }
  return (
    cell.value !== null &&
    cell.value !== undefined &&
    typeof cell.value === "object" &&
    !(cell.value instanceof Date)
  );
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
  const paths = originalWorksheets.map(
    (worksheet, index) => worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`,
  );
  const originalByName = new Map(
    originalWorksheets.map((worksheet) => [worksheet.name, worksheet]),
  );
  const originalDefinitionByPath = new Map<string, SheetDefinition>();
  for (const definition of originalDefinitions) {
    const original = originalByName.get(definition.name);
    if (!original) continue;
    const originalIndex = originalWorksheets.indexOf(original);
    const sourcePath = original.sourcePath ?? `xl/worksheets/sheet${originalIndex + 1}.xml`;
    originalDefinitionByPath.set(sourcePath.toLowerCase(), definition);
  }
  const originalPaths = new Set(paths.map((path) => path.toLowerCase()));
  const currentWorksheetByPath = new Map(
    worksheets.map((worksheet, index) => [
      (worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`).toLowerCase(),
      worksheet,
    ]),
  );
  const worksheetDefinitions = worksheets.map((worksheet, index): SheetDefinition => {
    const sourcePath = worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`;
    const definition = originalDefinitionByPath.get(sourcePath.toLowerCase());
    if (definition) {
      return {
        ...definition,
        name: worksheet.name ?? definition.name,
        state: worksheet.state ?? definition.state,
      };
    }
    return {
      name: worksheet.name ?? `Sheet${index + 1}`,
      sheetId: worksheet.sheetId,
      rId: "",
      state: worksheet.state,
    };
  });
  const existingWorksheetDefinitions = worksheets.flatMap((worksheet, index) => {
    const path = worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`;
    return originalPaths.has(path.toLowerCase()) ? [worksheetDefinitions[index]!] : [];
  });
  const newWorksheetDefinitions = worksheets.flatMap((worksheet, index) => {
    const path = worksheet.sourcePath ?? `xl/worksheets/sheet${index + 1}.xml`;
    return originalPaths.has(path.toLowerCase()) ? [] : [worksheetDefinitions[index]!];
  });
  let worksheetDefinitionIndex = 0;
  const sheetDefinitions: SheetDefinition[] = originalDefinitions.flatMap((definition) => {
    const original = originalByName.get(definition.name);
    if (!original) return [definition];
    const originalIndex = originalWorksheets.indexOf(original);
    const path = (
      original.sourcePath ?? `xl/worksheets/sheet${originalIndex + 1}.xml`
    ).toLowerCase();
    if (!currentWorksheetByPath.has(path)) return [];
    const replacement = existingWorksheetDefinitions[worksheetDefinitionIndex];
    worksheetDefinitionIndex += 1;
    return replacement ? [replacement] : [];
  });
  sheetDefinitions.push(...newWorksheetDefinitions);
  const usedSheetIds = new Set<number>();
  let nextSheetId = Math.max(
    0,
    ...originalDefinitions.map((definition) => definition.sheetId ?? 0),
  );
  let nextRelationship = Math.max(
    0,
    ...originalDefinitions.map((definition) => Number(definition.rId.replace(/\D/g, "") || 0)),
  );

  for (const definition of sheetDefinitions) {
    if (definition.sheetId === undefined) continue;
    if (usedSheetIds.has(definition.sheetId))
      throw new Error("Duplicate workbook sheet ID: " + definition.sheetId);
    usedSheetIds.add(definition.sheetId);
  }

  for (const [worksheetIndex, worksheet] of worksheets.entries()) {
    const sourcePath = worksheet.sourcePath ?? `xl/worksheets/sheet${worksheetIndex + 1}.xml`;
    if (originalPaths.has(sourcePath.toLowerCase())) continue;
    const definition = worksheetDefinitions[worksheetIndex]!;
    const existingId = definition.sheetId;
    const sheetId =
      existingId ??
      (() => {
        while (usedSheetIds.has(nextSheetId)) nextSheetId += 1;
        return nextSheetId;
      })();
    usedSheetIds.add(sheetId);
    nextRelationship += 1;
    const position = sheetDefinitions.indexOf(definition);
    sheetDefinitions[position] = {
      ...definition,
      sheetId,
      rId: `rId${nextRelationship}`,
    };
  }

  options.sheetDefinitions = sheetDefinitions;
}
