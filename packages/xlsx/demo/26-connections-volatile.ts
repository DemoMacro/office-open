// Connections, query table refresh, volatile dependencies, web publish objects.

import { writeFileSync } from "node:fs";

import type { Context } from "@office-open/core";
import { Workbook, Packer, QueryTableXml, ConnectionsXml } from "@office-open/xlsx";

const wb = new Workbook({
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
  volTypes: [
    {
      type: "realTimeData",
      mains: [
        {
          first: "rId1",
          topics: [
            {
              value: "StockPrice",
              stringTopics: ["Topic1"],
              refs: [{ reference: "A1", sheetIndex: 0 }],
            },
          ],
        },
      ],
    },
  ],
  webPublishObjects: [
    {
      rId: "rId1",
      destinationFile: "report.htm",
      title: "Price List",
      autoRepublish: true,
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);

// ── Standalone XML demos ──

const context: Context = { fileData: wb, stack: [] };

// Connections (xl/connections.xml)
const connections = new ConnectionsXml([
  {
    id: 1,
    name: "NorthwindDB",
    type: 3,
    refreshedVersion: 6,
    interval: 300,
    keepAlive: true,
    dbPr: {
      connection: "Provider=SQLOLEDB.1;Data Source=server;Initial Catalog=Northwind",
      command: "SELECT * FROM Products",
      commandType: 2,
    },
    parameters: [{ name: "MinPrice", parameterType: "value", integerValue: 10, integer: true }],
  },
  {
    id: 2,
    name: "WebQuery",
    type: 4,
    refreshedVersion: 6,
    webPr: {
      url: "https://example.com/data.csv",
      htmlFormat: "rtf",
      firstRow: true,
      textFields: [
        { type: 1, dataType: "text" },
        { type: 2, dataType: "text" },
      ],
    },
  },
]);
console.log("Connections XML:", connections.toXml(context).slice(0, 300) + "...");

// QueryTable with refresh options
const qt = new QueryTableXml({
  name: "QueryTable1",
  connectionId: 1,
  refreshOnLoad: true,
  queryTableRefresh: {
    nextId: 100,
    preserveFormatting: true,
    preserveSortFilterLayout: true,
    queryTableFields: [
      { id: 1, name: "Product" },
      { id: 2, name: "Price", numberFormatting: true, dataBound: true },
    ],
  },
});
console.log("QueryTable XML:", qt.toXml(context).slice(0, 300) + "...");
