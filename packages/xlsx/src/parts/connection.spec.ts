import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { connectionsDesc } from "./connection";
import { queryTableDesc } from "./query-table";

// ── Minimal context stubs (see external-link.spec.ts) ──

const writeCtx = {} as unknown as WriteContext;

const readCtx = {} as unknown as ReadContext;

function parseRoot(xml: string | undefined) {
  if (!xml) throw new Error("stringify produced no XML");
  const doc = parseXml(xml, { nativeTypeAttributes: true });
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return el;
}

describe("connectionsDesc", () => {
  it("round-trips a database connection with parameters", () => {
    const opts = {
      connections: [
        {
          id: 1,
          name: "DB",
          type: 2,
          refreshedVersion: 6,
          keepAlive: true,
          interval: 5,
          savePassword: true,
          credentials: "stored",
          dbPr: {
            connection: "Provider=SQLOLEDB;Data Source=db",
            command: "SELECT * FROM t",
            serverCommand: "SELECT * FROM v",
            commandType: 2,
          },
          parameters: [
            {
              name: "p1",
              sqlType: 5,
              parameterType: "cell" as const,
              refreshOnChange: true,
              prompt: "pick a cell",
              cell: "Sheet1!$A$1",
            },
          ],
        },
      ],
    };
    const xml = connectionsDesc.stringify(opts, writeCtx);
    const parsed = connectionsDesc.parse(parseRoot(xml), readCtx);
    expect(parsed.connections).toHaveLength(1);
    const c = parsed.connections[0]!;
    expect(c).toMatchObject({
      id: 1,
      name: "DB",
      type: 2,
      refreshedVersion: 6,
      keepAlive: true,
      interval: 5,
      savePassword: true,
      credentials: "stored",
    });
    expect(c.dbPr).toMatchObject({
      connection: "Provider=SQLOLEDB;Data Source=db",
      command: "SELECT * FROM t",
      serverCommand: "SELECT * FROM v",
      commandType: 2,
    });
    expect(c.parameters![0]).toMatchObject({
      name: "p1",
      sqlType: 5,
      parameterType: "cell",
      refreshOnChange: true,
      cell: "Sheet1!$A$1",
    });
  });

  it("round-trips a web query connection with table selections", () => {
    const opts = {
      connections: [
        {
          id: 2,
          name: "Web",
          type: 4,
          webPr: {
            url: "https://example.com/table",
            firstRow: true,
            htmlTables: true,
            htmlFormat: "rtf",
            tables: ["table1", 3, null] as (string | number | null)[],
          },
        },
      ],
    };
    const xml = connectionsDesc.stringify(opts, writeCtx);
    const parsed = connectionsDesc.parse(parseRoot(xml), readCtx);
    const w = parsed.connections[0]!.webPr!;
    expect(w.url).toBe("https://example.com/table");
    expect(w.firstRow).toBe(true);
    expect(w.htmlTables).toBe(true);
    expect(w.htmlFormat).toBe("rtf");
    expect(w.tables).toEqual(["table1", 3, null]);
  });

  it("round-trips a text-import connection with field layout", () => {
    const opts = {
      connections: [
        {
          id: 3,
          name: "Text",
          type: 5,
          textPr: {
            fileType: "dos" as const,
            codePage: 65001,
            firstRow: 2,
            sourceFile: "C:\\data\\in.csv",
            comma: true,
            qualifier: "doubleQuote" as const,
            textFields: [{ type: "text", position: 1 }],
          },
        },
      ],
    };
    const xml = connectionsDesc.stringify(opts, writeCtx);
    const parsed = connectionsDesc.parse(parseRoot(xml), readCtx);
    const t = parsed.connections[0]!.textPr!;
    expect(t.fileType).toBe("dos");
    expect(t.codePage).toBe(65001);
    expect(t.firstRow).toBe(2);
    expect(t.sourceFile).toBe("C:\\data\\in.csv");
    expect(t.comma).toBe(true);
    expect(t.qualifier).toBe("doubleQuote");
    expect(t.textFields).toEqual([{ type: "text", position: 1 }]);
  });
});

describe("queryTableDesc", () => {
  it("round-trips query table attributes and refresh layout", () => {
    const opts = {
      name: "QT",
      headers: false,
      rowNumbers: true,
      growShrinkType: "insertWhole" as const,
      fillFormulas: true,
      connectionId: 7,
      autoFormat: true,
      queryTableRefresh: {
        preserveSortFilterLayout: false,
        fieldIdWrapped: true,
        minimumVersion: 3,
        nextId: 12,
        queryTableFields: [
          { id: 1, name: "col1", rowNumbers: true, tableColumnId: 4 },
          { id: 2, name: "col2", dataBound: false },
        ],
        deletedFields: [{ name: "gone" }],
      },
    };
    const xml = queryTableDesc.stringify(opts, writeCtx);
    const parsed = queryTableDesc.parse(parseRoot(xml), readCtx);
    expect(parsed).toMatchObject({
      name: "QT",
      headers: false,
      rowNumbers: true,
      growShrinkType: "insertWhole",
      fillFormulas: true,
      connectionId: 7,
      autoFormat: true,
    });
    const r = parsed.queryTableRefresh!;
    expect(r.preserveSortFilterLayout).toBe(false);
    expect(r.fieldIdWrapped).toBe(true);
    expect(r.minimumVersion).toBe(3);
    expect(r.nextId).toBe(12);
    expect(r.queryTableFields![0]).toMatchObject({
      id: 1,
      name: "col1",
      rowNumbers: true,
      tableColumnId: 4,
    });
    expect(r.queryTableFields![1]).toMatchObject({ id: 2, dataBound: false });
    expect(r.deletedFields).toEqual([{ name: "gone" }]);
  });
});
