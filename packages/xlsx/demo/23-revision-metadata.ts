// Shared-workbook revision log.
//
// Generates xl/revisionHeaders.xml + xl/revisions/revisionN.xml + xl/users.xml,
// the three parts that record tracked changes in a shared workbook.
// Reference: OOXML transitional sml.xsd — CT_RevisionHeaders / CT_Revisions / CT_Users.

import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Data",
      rows: [
        { cells: [{ value: "Product" }, { value: "Price" }] },
        { cells: [{ value: "Widget" }, { value: 9.99 }] },
        { cells: [{ value: "Gadget" }, { value: 19.99 }] },
      ],
    },
  ],
  revisionLog: {
    // xl/revisionHeaders.xml — index of editing sessions.
    headers: {
      guid: "{A84A6777-8908-4CB9-9EB6-625CEFF419D3}",
      revisionId: 1,
      version: 2,
      diskRevisions: true,
      headers: [
        {
          guid: "{CFEA9B63-728B-4274-A346-0440E1573AB4}",
          dateTime: "2026-06-19T10:00:00Z",
          userName: "Alice",
          // rId1 → xl/revisions/revision1.xml (matches logs[0]).
          rId: "rId1",
          maxSheetId: 1,
          sheetIds: [1],
        },
      ],
    },
    // xl/revisions/revision1.xml — the concrete changes in this session.
    logs: [
      {
        revisions: [
          // Cell change: A1 set to "foo" (nc = new cell, raw CT_Cell XML).
          {
            type: "cellChange",
            data: {
              rId: 1,
              sheetId: 1,
              newCellXml: `<nc r="A1" t="inlineStr"><is><t>foo</t></is></nc>`,
            },
          },
          // Comment added on B2 (CT_RevisionComment has no AG_RevData, so no rId).
          {
            type: "comment",
            data: {
              sheetId: 1,
              cell: "B2",
              guid: "{841DBE00-ECD0-478E-893B-30CE5DABBEF5}",
              action: "add",
              author: "Alice",
              newLength: 5,
            },
          },
        ],
      },
    ],
    // xl/users.xml — who is editing the shared workbook.
    users: {
      users: [
        {
          guid: "{11111111-2222-3333-4444-555555555555}",
          name: "Alice",
          id: 1,
          dateTime: "2026-06-19T10:00:00Z",
        },
      ],
    },
  },
});

writeFileSync("My Workbook.xlsx", buffer);
