import { unzipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./generate";
import { parseWorkbookSync } from "./parse";
import type { WorkbookOptions } from "./parts/file";
import type { CommentOptions } from "./parts/worksheet/types";
import { patchWorkbook } from "./patch";

const workbook = (): WorkbookOptions => ({
  title: "Original",
  sharedStringsCount: 2,
  sharedStringsUniqueCount: 2,
  worksheets: [
    {
      name: "Data",
      rows: [
        {
          cells: [
            { reference: "A1", value: "Hello {{name}}" },
            { reference: "B1", value: "Shared" },
          ],
        },
      ],
    },
  ],
});

const source = async (): Promise<Uint8Array> =>
  (await generateWorkbook(workbook(), { type: "uint8array" })) as Uint8Array;

const twoSheetSource = async (): Promise<Uint8Array> =>
  (await generateWorkbook({ worksheets: [{ name: "A" }, { name: "B" }] })) as Uint8Array;

describe("patchWorkbook commands", () => {
  it("assigns unique paths and identities to worksheets inserted after removal", async () => {
    const patched = (await patchWorkbook({
      data: await twoSheetSource(),
      commands: [
        { type: "remove", at: { field: "worksheets", route: [{ name: "A" }] } },
        {
          type: "insert",
          at: { field: "worksheets" },
          position: "end",
          value: [{ name: "C", rows: [] }],
        },
      ],
    })) as Uint8Array;
    const archive = unzipSync(patched);
    const parsed = parseWorkbookSync(patched);

    expect(parsed.worksheets?.map((worksheet) => worksheet.sourcePath)).toEqual([
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/sheet3.xml",
    ]);
    expect(
      new Set(Object.keys(archive).filter((path) => path.startsWith("xl/worksheets/"))),
    ).toEqual(new Set(["xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]));
    expect(new Set(parsed.sheetDefinitions?.map((sheet) => sheet.sheetId))).toHaveLength(2);
    expect(new Set(parsed.sheetDefinitions?.map((sheet) => sheet.rId))).toHaveLength(2);
  });

  it("preserves an inserted worksheet's explicit sheet id", async () => {
    const parsed = parseWorkbookSync(
      (await patchWorkbook({
        data: await twoSheetSource(),
        commands: [
          {
            type: "insert",
            at: { field: "worksheets" },
            position: "end",
            value: [{ name: "C", rows: [], sheetId: 99 }],
          },
        ],
      })) as Uint8Array,
    );

    expect(parsed.sheetDefinitions?.at(-1)).toMatchObject({ name: "C", sheetId: 99 });
    expect(new Set(parsed.sheetDefinitions?.map((sheet) => sheet.sheetId))).toHaveLength(3);
  });

  it("removes the content-type override for a deleted worksheet", async () => {
    const patched = (await patchWorkbook({
      data: await twoSheetSource(),
      commands: [{ type: "remove", at: { field: "worksheets", route: [{ name: "B" }] } }],
    })) as Uint8Array;
    const archive = unzipSync(patched);
    const contentTypes = new TextDecoder().decode(archive["[Content_Types].xml"]!);

    expect(archive["xl/worksheets/sheet2.xml"]).toBeUndefined();
    expect(contentTypes).not.toContain('PartName="/xl/worksheets/sheet2.xml"');
  });

  it("replaces addressed canonical text", async () => {
    const parsed = parseWorkbookSync(
      await patchWorkbook({
        data: await source(),
        commands: [
          {
            type: "replace",
            at: {
              field: "worksheets",
              route: [{ name: "Data" }],
              text: { placeholder: "name" },
            },
            value: "Alice",
          },
        ],
      }),
    );

    const patchedArchive = unzipSync(
      (await patchWorkbook({
        data: await source(),
        commands: [
          {
            type: "replace",
            at: {
              field: "worksheets",
              route: [{ name: "Data" }],
              text: { placeholder: "name" },
            },
            value: "Alice",
          },
        ],
      })) as Uint8Array,
    );
    const sharedStrings = new TextDecoder().decode(patchedArchive["xl/sharedStrings.xml"]!);
    expect(sharedStrings).toContain('count="2"');
    expect(sharedStrings).toContain('uniqueCount="2"');
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toBe("Hello Alice");
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.sharedIndex).toBe(0);
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[1]).toMatchObject({
      sharedIndex: 1,
      value: "Shared",
    });
  });

  it("deduplicates shared strings after multiple replacements", async () => {
    const cellRoute = (index: number) => [
      { name: "Data" },
      { collection: "rows" },
      { index: 0 },
      { collection: "cells" },
      { index },
    ];
    const sameValue = async (): Promise<Uint8Array> =>
      (await generateWorkbook({
        sharedStringsCount: 2,
        sharedStringsUniqueCount: 1,
        worksheets: [
          {
            name: "Data",
            rows: [
              {
                cells: [
                  { reference: "A1", value: "Old" },
                  { reference: "B1", value: "Old" },
                ],
              },
            ],
          },
        ],
      })) as Uint8Array;
    const split = parseWorkbookSync(
      await patchWorkbook({
        data: await sameValue(),
        commands: [
          {
            type: "replace",
            at: {
              field: "worksheets",
              route: cellRoute(0),
            },
            value: { reference: "A1", value: "First" },
          },
          {
            type: "replace",
            at: {
              field: "worksheets",
              route: cellRoute(1),
            },
            value: { reference: "B1", value: "Second" },
          },
        ],
      }),
    );
    const splitCells = split.worksheets?.[0]?.rows?.[0]?.cells ?? [];
    expect(splitCells.map((cell) => cell.value)).toEqual(["First", "Second"]);
    expect(new Set(splitCells.map((cell) => cell.sharedIndex))).toHaveLength(2);
    const merged = parseWorkbookSync(
      (await patchWorkbook({
        data: await sameValue(),
        commands: [
          {
            type: "replace",
            at: {
              field: "worksheets",
              route: cellRoute(0),
            },
            value: { reference: "A1", value: "New" },
          },
          {
            type: "replace",
            at: {
              field: "worksheets",
              route: cellRoute(1),
            },
            value: { reference: "B1", value: "New" },
          },
        ],
      })) as Uint8Array,
    );
    const mergedCells = merged.worksheets?.[0]?.rows?.[0]?.cells ?? [];
    expect(mergedCells.map((cell) => cell.value)).toEqual(["New", "New"]);
    expect(new Set(mergedCells.map((cell) => cell.sharedIndex))).toHaveLength(1);
    expect(merged.sharedStrings).toEqual(["New"]);
    expect(merged.sharedStringsUniqueCount).toBe(1);
    expect(
      new TextDecoder().decode(
        unzipSync((await patchWorkbook({ data: await sameValue(), commands: [] })) as Uint8Array)[
          "xl/sharedStrings.xml"
        ]!,
      ),
    ).toContain('uniqueCount="1"');
  });

  it("updates a worksheet and root metadata", async () => {
    const parsed = parseWorkbookSync(
      await patchWorkbook({
        data: await source(),
        commands: [
          {
            type: "update",
            at: { field: "worksheets", route: [{ name: "Data" }] },
            set: { state: "hidden" },
          },
          { type: "update", at: { field: "title" }, set: "Patched" },
        ],
      }),
    );

    expect(parsed.worksheets?.[0]?.state).toBe("hidden");
    expect(parsed.title).toBe("Patched");
  });

  it("inserts worksheet comments into an absent collection", async () => {
    const comments: CommentOptions[] = [{ cell: "A1", author: "Alice", text: "needs review" }];
    const parsed = parseWorkbookSync(
      await patchWorkbook({
        data: await source(),
        commands: [
          {
            type: "insert",
            at: { field: "worksheets", route: [{ name: "Data" }, { collection: "comments" }] },
            value: comments,
          },
        ],
      }),
    );

    expect(parsed.worksheets?.[0]?.comments).toHaveLength(1);
  });

  it("moves and removes worksheets in command order", async () => {
    const options = workbook();
    options.worksheets?.push({ name: "Summary", rows: [] }, { name: "Archive", rows: [] });
    const parsed = parseWorkbookSync(
      await patchWorkbook({
        data: await generateWorkbook(options, { type: "uint8array" }),
        commands: [
          { type: "move", at: { field: "worksheets", route: [{ name: "Summary" }] }, to: "start" },
          { type: "remove", at: { field: "worksheets", route: [{ name: "Archive" }] } },
        ],
      }),
    );

    expect(parsed.worksheets?.map((worksheet) => worksheet.name)).toEqual(["Summary", "Data"]);
  });

  it("rejects a missing route without emitting output", async () => {
    await expect(
      patchWorkbook({
        data: await source(),
        commands: [{ type: "remove", at: { field: "worksheets", route: [{ name: "Missing" }] } }],
      }),
    ).rejects.toThrow(/Missing/);
  });
});
