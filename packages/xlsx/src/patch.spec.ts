import { unzipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./generate";
import { parseWorkbookSync } from "./parse";
import type { WorkbookOptions } from "./parts/file";
import type { CommentOptions } from "./parts/worksheet/types";
import { patchWorkbook } from "./patch";

const workbook = (): WorkbookOptions => ({
  title: "Original",
  worksheets: [
    {
      name: "Data",
      rows: [{ cells: [{ reference: "A1", value: "Hello {{name}}" }] }],
    },
  ],
});

const source = async (): Promise<Uint8Array> =>
  (await generateWorkbook(workbook(), { type: "uint8array" })) as Uint8Array;

const twoSheetSource = async (): Promise<Uint8Array> =>
  (await generateWorkbook({ worksheets: [{ name: "A" }, { name: "B" }] })) as Uint8Array;

describe("patchWorkbook commands", () => {
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

    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toBe("Hello Alice");
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
