// XML mapping: MapInfo for mapping XML schema to spreadsheet cells.

import type { Context } from "@office-open/core";
import { MapInfoXml, SingleXmlCellsXml, XmlColumnPrXml } from "@office-open/xlsx";

const context: Context = { fileData: null as never, stack: [] };

// MapInfo (xl/xmlMaps.xml) — maps XML schema elements to worksheet cells
const mapInfo = new MapInfoXml({
  selectionNamespaces: "xmlns:ord='urn:example:orders'",
  schemas: [
    {
      id: "Schema1",
      namespace: "urn:example:orders",
      schemaRef: "orders.xsd",
    },
  ],
  maps: [
    {
      id: 1,
      name: "Order_Map",
      rootElement: "order",
      schemaID: "Schema1",
      dataBinding: {
        dataBindingName: "Order_Binding",
        connectionID: 1,
      },
    },
  ],
});
console.log("MapInfo XML:", mapInfo.toXml(context).slice(0, 400) + "...");

// SingleXmlCells — mapping single cells to XML elements
const sxc = new SingleXmlCellsXml([
  {
    id: 1,
    r: "A1",
    connectionId: 1,
    xmlCellPr: {
      id: 1,
      xmlPr: { mapId: 1, xpath: "/order/product", xmlDataType: "string" },
    },
  },
]);
console.log("SingleXmlCells XML:", sxc.toXml(context).slice(0, 300) + "...");

// XmlColumnPr — column-level XML mapping
const xcp = new XmlColumnPrXml({
  xpath: "/order/quantity",
  xmlDataType: "integer",
  mapId: 1,
});
console.log("XmlColumnPr XML:", xcp.toXml(context).slice(0, 200) + "...");
