import { unzipSync, zipSync } from "@office-open/core";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "../generate";
import { parseWorkbookSync } from "../parse";
import type { WorkbookOptions } from "./file";
import { revisionHeadersDesc, revisionLogDesc, usersDesc } from "./revision-log";
import type {
  RevisionHeadersOptions,
  RevisionLogOptions,
  RevisionNestedChild,
  UsersOptions,
} from "./revision-log";

// revision descriptors ignore their context, so empty stubs suffice.
const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function parseRoot(xml: string) {
  // nativeTypeAttributes mirrors the real xlsx parse path (ParsedArchive.get
  // coerces "1"/"0" to numbers), so boolean reads are exercised under coercion.
  const el = parseXml(xml, { nativeTypeAttributes: true }).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return el;
}

describe("revisionHeadersDesc round-trip", () => {
  it("round-trips headers with one entry and root flags", () => {
    const opts: RevisionHeadersOptions = {
      guid: "{AAA}",
      revisionId: 1,
      version: 2,
      diskRevisions: true,
      trackRevisions: true,
      headers: [
        {
          guid: "{BBB}",
          dateTime: "2026-06-19T10:00:00Z",
          userName: "Alice",
          rId: "rId1",
          maxSheetId: 1,
          sheetIds: [1],
        },
      ],
    };
    const result = revisionHeadersDesc.parse(
      parseRoot(revisionHeadersDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    expect(result.guid).toBe("{AAA}");
    expect(result.revisionId).toBe(1);
    expect(result.diskRevisions).toBe(true);
    expect(result.headers).toHaveLength(1);
    expect(result.headers[0]).toMatchObject({
      guid: "{BBB}",
      userName: "Alice",
      rId: "rId1",
      maxSheetId: 1,
      sheetIds: [1],
    });
  });

  it("round-trips boolean root flags under nativeTypeAttributes coercion", () => {
    const opts: RevisionHeadersOptions = {
      guid: "{A}",
      diskRevisions: true,
      trackRevisions: true,
      history: true,
      keepChangeHistory: true,
      protected: true,
      headers: [
        { guid: "{B}", dateTime: "t", userName: "U", rId: "rId1", maxSheetId: 1, sheetIds: [1] },
      ],
    };
    const result = revisionHeadersDesc.parse(
      parseRoot(revisionHeadersDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    // "1" coerces to number 1 on the real parse path; String() guard recovers it
    expect(result.diskRevisions).toBe(true);
    expect(result.trackRevisions).toBe(true);
    expect(result.history).toBe(true);
    expect(result.keepChangeHistory).toBe(true);
    expect(result.protected).toBe(true);
  });

  it("round-trips reviewed revision ids", () => {
    const opts: RevisionHeadersOptions = {
      guid: "{A}",
      headers: [
        {
          guid: "{B}",
          dateTime: "t",
          userName: "U",
          rId: "rId1",
          maxSheetId: 1,
          sheetIds: [1],
          reviewed: [3, 7],
        },
      ],
    };
    const result = revisionHeadersDesc.parse(
      parseRoot(revisionHeadersDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    expect(result.headers[0]?.reviewed).toEqual([3, 7]);
  });
});

describe("usersDesc round-trip", () => {
  it("empty users returns undefined from stringify", () => {
    expect(usersDesc.stringify({ users: [] }, writeCtx)).toBeUndefined();
  });

  it("round-trips user entries", () => {
    const opts: UsersOptions = {
      users: [{ guid: "{U1}", name: "Alice", id: 1, dateTime: "2026-06-19T10:00:00Z" }],
    };
    const result = usersDesc.parse(parseRoot(usersDesc.stringify(opts, writeCtx)!), readCtx);
    expect(result.users).toHaveLength(1);
    expect(result.users![0]).toEqual({
      guid: "{U1}",
      name: "Alice",
      id: 1,
      dateTime: "2026-06-19T10:00:00Z",
    });
  });
});

describe("revisionLogDesc round-trip", () => {
  it("empty revisions returns undefined from stringify", () => {
    expect(revisionLogDesc.stringify({ revisions: [] }, writeCtx)).toBeUndefined();
  });

  it("round-trips cellChange with raw new-cell XML", () => {
    const opts: RevisionLogOptions = {
      revisions: [
        {
          type: "cellChange",
          data: {
            rId: 1,
            sheetId: 1,
            newCellXml: `<nc r="A1" t="inlineStr"><is><t>foo</t></is></nc>`,
          },
        },
      ],
    };
    const result = revisionLogDesc.parse(
      parseRoot(revisionLogDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    expect(result.revisions).toHaveLength(1);
    expect(result.revisions[0]?.type).toBe("cellChange");
    const d = result.revisions[0]?.data as { rId: number; sheetId: number; newCellXml: string };
    expect(d.rId).toBe(1);
    expect(d.sheetId).toBe(1);
    expect(d.newCellXml).toContain("<t>foo</t>");
  });

  it("round-trips complete cell change payloads in XSD order", () => {
    const oldCellXml = `<oc r="A1" t="inlineStr"><is><t>old</t></is></oc>`;
    const newCellXml = `<nc r="A1" t="inlineStr"><is><t>new</t></is></nc>`;
    const oldDxfXml = `<odxf><font><b/></font></odxf>`;
    const newDxfXml = `<ndxf><font><i/></font></ndxf>`;
    const opts: RevisionLogOptions = {
      revisions: [
        {
          type: "cellChange",
          data: {
            rId: 1,
            sheetId: 1,
            oldCellXml,
            newCellXml,
            oldDxfXml,
            newDxfXml,
            hasOldDxf: true,
            xfDxf: true,
          },
        },
      ],
    };
    const result = revisionLogDesc.parse(
      parseRoot(revisionLogDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    const xml = revisionLogDesc.stringify(opts, writeCtx)!;
    expect(xml.indexOf("<oc")).toBeLessThan(xml.indexOf("<nc"));
    expect(xml.indexOf("<nc")).toBeLessThan(xml.indexOf("<odxf"));
    expect(xml.indexOf("<odxf")).toBeLessThan(xml.indexOf("<ndxf"));
    const data = result.revisions[0]?.data as {
      oldCellXml?: string;
      newCellXml?: string;
      oldDxfXml?: string;
      newDxfXml?: string;
    };
    expect(data.oldCellXml).toBe(oldCellXml);
    expect(data.newCellXml).toBe(newCellXml);
    expect(data.oldDxfXml).toBe(oldDxfXml);
    expect(data.newDxfXml).toBe(newDxfXml);
  });

  it("round-trips rowColumn with structured undo/rcc/rfmt children", () => {
    const opts: RevisionLogOptions = {
      revisions: [
        {
          type: "rowColumn",
          data: {
            rId: 3,
            sheetId: 1,
            ref: "2:2",
            action: "insertRow",
            children: [
              {
                kind: "undo",
                data: {
                  index: 1,
                  expression: "ref",
                  undoRange: "A2:B2",
                  ref3D: true,
                  numberFormat: true,
                },
              },
              {
                kind: "cellChange",
                data: { rId: 4, sheetId: 1, newCellXml: `<nc r="A2"><v>7</v></nc>` },
              },
              { kind: "formatting", data: { sheetId: 1, sqref: "A2" } },
            ],
          },
        },
      ],
    };
    const result = revisionLogDesc.parse(
      parseRoot(revisionLogDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    const d = result.revisions[0]?.data as { children?: RevisionNestedChild[] };
    expect(d.children).toHaveLength(3);
    expect(d.children![0]).toMatchObject({
      kind: "undo",
      data: { index: 1, expression: "ref", undoRange: "A2:B2", ref3D: true, numberFormat: true },
    });
    expect(d.children![1]?.kind).toBe("cellChange");
    expect(d.children![2]?.kind).toBe("formatting");
  });

  it("round-trips comment (no AG_RevData, no rId)", () => {
    const opts: RevisionLogOptions = {
      revisions: [
        {
          type: "comment",
          data: {
            sheetId: 1,
            cell: "B2",
            guid: "{C}",
            action: "add",
            author: "Alice",
            newLength: 5,
          },
        },
      ],
    };
    const result = revisionLogDesc.parse(
      parseRoot(revisionLogDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    expect(result.revisions[0]).toMatchObject({
      type: "comment",
      data: { sheetId: 1, cell: "B2", guid: "{C}", action: "add", author: "Alice", newLength: 5 },
    });
  });

  it("round-trips insertSheet, sheetRename, customView, conflict, queryTableField", () => {
    const opts: RevisionLogOptions = {
      revisions: [
        { type: "insertSheet", data: { rId: 2, sheetId: 4, name: "Sheet4", sheetPosition: 3 } },
        { type: "sheetRename", data: { rId: 3, sheetId: 1, oldName: "Old", newName: "New" } },
        { type: "customView", data: { guid: "{V}", action: "add" } },
        { type: "conflict", data: { rId: 4, sheetId: 1, rejected: true } },
        { type: "queryTableField", data: { sheetId: 1, ref: "A1:A3", fieldId: 2 } },
      ],
    };
    const result = revisionLogDesc.parse(
      parseRoot(revisionLogDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    expect(result.revisions).toHaveLength(5);
    expect(result.revisions[0]).toMatchObject({
      type: "insertSheet",
      data: { sheetId: 4, name: "Sheet4", sheetPosition: 3 },
    });
    expect(result.revisions[1]).toMatchObject({
      type: "sheetRename",
      data: { oldName: "Old", newName: "New" },
    });
    expect(result.revisions[2]).toMatchObject({
      type: "customView",
      data: { guid: "{V}", action: "add" },
    });
    expect(result.revisions[3]).toMatchObject({
      type: "conflict",
      data: { sheetId: 1, rejected: true },
    });
    expect(result.revisions[4]).toMatchObject({
      type: "queryTableField",
      data: { ref: "A1:A3", fieldId: 2 },
    });
  });

  it("round-trips definedName with the full attribute set", () => {
    const opts: RevisionLogOptions = {
      revisions: [
        {
          type: "definedName",
          data: {
            rId: 5,
            name: "MyRange",
            localSheetId: 0,
            hidden: true,
            function: true,
            comment: "a comment",
            formula: "Sheet1!$A$1",
          },
        },
      ],
    };
    const result = revisionLogDesc.parse(
      parseRoot(revisionLogDesc.stringify(opts, writeCtx)!),
      readCtx,
    );
    const entry = result.revisions[0];
    if (!entry) throw new Error("expected definedName revision");
    expect(entry.type).toBe("definedName");
    if (entry.type !== "definedName") return;
    const d = entry.data;
    expect(d.name).toBe("MyRange");
    expect(d.localSheetId).toBe(0);
    expect(d.hidden).toBe(true);
    expect(d.function).toBe(true);
    expect(d.comment).toBe("a comment");
    expect(d.formula).toBe("Sheet1!$A$1");
  });
});

const revisionOptions = (): WorkbookOptions => ({
  worksheets: [{ name: "Data", rows: [{ cells: [{ value: "Product" }] }] }],
  revisionLog: {
    headers: {
      guid: "{HDR}",
      headers: [
        {
          guid: "{H1}",
          dateTime: "2026-06-19T10:00:00Z",
          userName: "Alice",
          rId: "rId1",
          maxSheetId: 1,
          sheetIds: [1],
        },
      ],
    },
    logs: [
      {
        revisions: [
          {
            type: "cellChange",
            data: {
              rId: 1,
              sheetId: 1,
              newCellXml: `<nc r="A1" t="inlineStr"><is><t>foo</t></is></nc>`,
            },
          },
        ],
      },
    ],
    users: { users: [{ guid: "{U}", name: "Alice", id: 1, dateTime: "2026-06-19T10:00:00Z" }] },
  },
});

describe("revision end-to-end round-trip", () => {
  it("generate → parse preserves revisionLog", async () => {
    const buffer = await generateWorkbook({
      worksheets: [{ name: "Data", rows: [{ cells: [{ value: "Product" }] }] }],
      revisionLog: {
        headers: {
          guid: "{HDR}",
          revisionId: 1,
          version: 2,
          headers: [
            {
              guid: "{H1}",
              dateTime: "2026-06-19T10:00:00Z",
              userName: "Alice",
              rId: "rId1",
              maxSheetId: 1,
              sheetIds: [1],
            },
          ],
        },
        logs: [
          {
            revisions: [
              {
                type: "cellChange",
                data: {
                  rId: 1,
                  sheetId: 1,
                  newCellXml: `<nc r="A1" t="inlineStr"><is><t>foo</t></is></nc>`,
                },
              },
              {
                type: "comment",
                data: {
                  sheetId: 1,
                  cell: "B2",
                  guid: "{CMT}",
                  action: "add",
                  author: "Alice",
                  newLength: 5,
                },
              },
            ],
          },
        ],
        users: { users: [{ guid: "{U}", name: "Alice", id: 1, dateTime: "2026-06-19T10:00:00Z" }] },
      },
    });
    const parsed = parseWorkbookSync(buffer);
    expect(parsed.revisionLog).toBeDefined();
    expect(parsed.revisionLog!.headers.headers[0]?.userName).toBe("Alice");
    expect(parsed.revisionLog!.headers.headers[0]?.rId).toBe("rId1");
    expect(parsed.revisionLog!.logs).toHaveLength(1);
    expect(parsed.revisionLog!.logs[0]?.revisions).toHaveLength(2);
    expect(parsed.revisionLog!.logs[0]?.revisions[0]?.type).toBe("cellChange");
    expect(parsed.revisionLog!.logs[0]?.revisions[1]?.type).toBe("comment");
    expect(parsed.revisionLog!.users?.users?.[0]?.name).toBe("Alice");
  });

  it("preserves a nonstandard revision part topology", async () => {
    const source = (await generateWorkbook(revisionOptions(), {
      type: "uint8array",
    })) as Uint8Array;
    const archive = unzipSync(source);
    const move = (from: string, to: string): void => {
      archive[to] = archive[from]!;
      delete archive[from];
    };
    move("xl/revisionHeaders.xml", "xl/revisions/revisionHeaders.xml");
    move("xl/users.xml", "xl/revisions/users.xml");
    move("xl/_rels/revisionHeaders.xml.rels", "xl/revisions/_rels/revisionHeaders.xml.rels");
    const replace = (path: string, from: string, to: string): void => {
      archive[path] = new TextEncoder().encode(
        new TextDecoder().decode(archive[path]!).replaceAll(from, to),
      );
    };
    replace("xl/revisions/revisionHeaders.xml", 'r:id="rId1"', 'r:id="rId7"');
    replace(
      "xl/revisions/_rels/revisionHeaders.xml.rels",
      'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog" Target="revisions/revision1.xml"',
      'Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog" Target="revision1.xml"',
    );
    replace(
      "xl/_rels/workbook.xml.rels",
      'Target="revisionHeaders.xml"',
      'Target="revisions/revisionHeaders.xml"',
    );
    replace("xl/_rels/workbook.xml.rels", 'Target="users.xml"', 'Target="revisions/users.xml"');
    replace(
      "[Content_Types].xml",
      'PartName="/xl/revisionHeaders.xml"',
      'PartName="/xl/revisions/revisionHeaders.xml"',
    );
    replace(
      "[Content_Types].xml",
      'PartName="/xl/users.xml"',
      'PartName="/xl/revisions/users.xml"',
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.revisionLog?.headersPath).toBe("revisions/revisionHeaders.xml");
    expect(parsed.revisionLog?.usersPath).toBe("revisions/users.xml");
    expect(parsed.revisionLog?.logs[0]).toMatchObject({
      path: "revisions/revision1.xml",
      relationshipTarget: "revision1.xml",
    });
    expect(parsed.rawParts).toBeUndefined();

    const output = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const result = unzipSync(output);
    expect(result["xl/revisions/revisionHeaders.xml"]).toBeDefined();
    const headerRels = new TextDecoder().decode(
      result["xl/revisions/_rels/revisionHeaders.xml.rels"]!,
    );
    expect(headerRels).toContain('Id="rId7"');
    expect(headerRels).toContain('Target="revision1.xml"');
    expect(result["xl/revisions/users.xml"]).toBeDefined();
    expect(result["xl/revisionHeaders.xml"]).toBeUndefined();
    expect(result["xl/_rels/revisionHeaders.xml.rels"]).toBeUndefined();
    expect(new TextDecoder().decode(result["[Content_Types].xml"]!)).toContain(
      'PartName="/xl/revisions/revisionHeaders.xml"',
    );
  });

  it("preserves revision log content-type declarations", async () => {
    const source = (await generateWorkbook(revisionOptions(), {
      type: "uint8array",
    })) as Uint8Array;
    const archive = unzipSync(source);
    const contentType =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml";
    archive["[Content_Types].xml"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["[Content_Types].xml"]!)
        .replace(
          "</Types>",
          `<Override PartName="/xl/revisions/revisionLog11.xml" ContentType="${contentType}"/></Types>`,
        ),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.contentTypes?.overrides).toContainEqual({
      partName: "/xl/revisions/revisionLog11.xml",
      contentType,
    });

    const output = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    expect(new TextDecoder().decode(unzipSync(output)["[Content_Types].xml"]!)).toContain(
      'PartName="/xl/revisions/revisionLog11.xml"',
    );
  });
});
