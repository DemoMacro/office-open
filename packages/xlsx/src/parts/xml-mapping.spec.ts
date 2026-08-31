import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { mapInfoDesc, singleXmlCellsDesc } from "./xml-mapping";
import type { MapInfoOptions } from "./xml-mapping";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function parseRoot(xml: string | undefined) {
  if (!xml) throw new Error("stringify produced no XML");
  const doc = parseXml(xml, { nativeTypeAttributes: true });
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return el;
}

describe("mapInfoDesc", () => {
  it("round-trips schemas with raw content and maps with data bindings", () => {
    const opts: MapInfoOptions = {
      selectionNamespaces: 'xmlns:ns1="http://example.com/catalog"',
      schemas: [
        {
          id: "Schema1",
          namespace: "http://example.com/catalog",
          schemaLanguage: "http://www.w3.org/2001/XMLSchema",
          content: '<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"/>',
        },
      ],
      maps: [
        {
          id: 1,
          name: "CatalogMap",
          rootElement: "catalog",
          schemaId: "Schema1",
          append: true,
          preserveFormat: true,
          dataBinding: {
            dataBindingName: "CatalogBinding",
            connectionId: 1,
            dataBindingLoadMode: 1,
          },
        },
      ],
    };
    const parsed = mapInfoDesc.parse(parseRoot(mapInfoDesc.stringify(opts, writeCtx)), readCtx);
    expect(parsed.selectionNamespaces).toBe('xmlns:ns1="http://example.com/catalog"');
    const schema = parsed.schemas[0]!;
    expect(schema.id).toBe("Schema1");
    expect(schema.namespace).toBe("http://example.com/catalog");
    expect(schema.schemaLanguage).toBe("http://www.w3.org/2001/XMLSchema");
    expect(schema.content).toContain("xsd:schema");
    const map = parsed.maps[0]!;
    expect(map).toMatchObject({
      id: 1,
      name: "CatalogMap",
      rootElement: "catalog",
      schemaId: "Schema1",
      append: true,
      preserveFormat: true,
    });
    expect(map.dataBinding).toMatchObject({
      dataBindingName: "CatalogBinding",
      connectionId: 1,
      dataBindingLoadMode: 1,
    });
  });
});

describe("singleXmlCellsDesc", () => {
  it("round-trips single-cell XML tables with xml cell properties", () => {
    const xml = singleXmlCellsDesc.stringify(
      {
        cells: [
          {
            id: 1,
            reference: "A1",
            connectionId: 2,
            properties: {
              id: 1,
              uniqueName: "PriceCell",
              mapping: { mapId: 1, xpath: "/catalog/item/price", xmlDataType: "double" },
            },
          },
        ],
      },
      writeCtx,
    );
    const parsed = singleXmlCellsDesc.parse(parseRoot(xml), readCtx);
    expect(parsed.cells).toHaveLength(1);
    expect(parsed.cells[0]).toMatchObject({ id: 1, reference: "A1", connectionId: 2 });
    expect(parsed.cells[0]!.properties).toMatchObject({ id: 1, uniqueName: "PriceCell" });
    expect(parsed.cells[0]!.properties.mapping).toEqual({
      mapId: 1,
      xpath: "/catalog/item/price",
      xmlDataType: "double",
    });
  });
});
