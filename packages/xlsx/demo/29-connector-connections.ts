import * as fs from "fs";

import { generateWorkbook } from "@office-open/xlsx";
import type { WorkbookOptions } from "@office-open/xlsx";

// Demonstrates connector non-visual properties: locking (a:cxnSpLocks) and
// endpoint connections (a:stCxn / a:endCxn) that glue a connector to shape
// connection sites. Round-trip preserves these so connectors stay attached
// when shapes move.

const options: WorkbookOptions = {
  worksheets: [
    {
      name: "Connections",
      rows: [{ cells: [{ value: "Connector connections" }] }],
      connectors: [
        {
          col: 2,
          row: 2,
          toCol: 8,
          toRow: 6,
          spPr: { geometry: "line" },
          locking: { noAdjustHandles: true, noChangeShapeType: true },
          startConnection: { id: 1, index: 0 },
          endConnection: { id: 2, index: 3 },
        },
      ],
    },
  ],
};

const buffer = await generateWorkbook(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/29-connector-connections.xlsx", buffer);
