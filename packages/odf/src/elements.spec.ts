import { parse } from "@office-open/xml";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { generateChartDocument, parseChartDocument } from "./chart";
import { generateDatabaseDocument, parseDatabaseDocument } from "./db";
import { ODF_ELEMENT_NAMES, parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
import { generateOdt, parseOdt } from "./odt";
import { generateOcf, manifestXml } from "./package";

describe("ODF element registry", () => {
  it("contains every unique element in the ODF and manifest schemas", () => {
    expect(ODF_ELEMENT_NAMES).toHaveLength(610);
    expect(new Set(ODF_ELEMENT_NAMES).size).toBe(610);
  });

  it("round-trips generic ODF nodes and preserves foreign namespaces", () => {
    const nodes: OdfXmlNode[] = [
      {
        name: "style:style",
        attributes: { "style:name": "List", "style:family": "paragraph" },
        children: [
          { name: "style:paragraph-properties", attributes: { "fo:text-align": "start" } },
        ],
      },
      { name: "text:list", children: [{ name: "text:list-item", children: ["Item"] }] },
      { name: "draw:frame", attributes: { "draw:name": "Frame" } },
      { name: "table:named-range", attributes: { "table:name": "Total" } },
      { name: "number:number-style", attributes: { "style:name": "N1" } },
      { name: "chart:chart", attributes: { "chart:class": "chart:bar" } },
      { name: "anim:par", children: [{ name: "anim:set" }] },
      { name: "config:config-item", attributes: { "config:name": "ShowGrid" } },
      { name: "db:data-source", attributes: { "db:name": "Library" } },
      { name: "form:form", attributes: { "form:name": "Editor" } },
      { name: "manifest:file-entry", attributes: { "manifest:full-path": "/" } },
      { name: "unknown:element", attributes: { hidden: "true" } },
    ];
    const serialized = serializeOdfNodes(nodes).join("");
    const document = parse(
      `<root xmlns:style="urn:1" xmlns:text="urn:2" xmlns:draw="urn:3" xmlns:table="urn:4" xmlns:number="urn:5" xmlns:chart="urn:6" xmlns:anim="urn:7" xmlns:config="urn:8" xmlns:db="urn:9" xmlns:form="urn:10" xmlns:manifest="urn:11" xmlns:unknown="urn:12">${serialized}</root>`,
      { ignoreDeclaration: true },
    );
    const parsed = parseOdfNodes(document.elements?.[0]);
    expect(parsed).toEqual(nodes);
  });

  it("round-trips a generic ODT extension element", () => {
    const extension: OdfXmlNode = {
      name: "text:bookmark",
      attributes: { "text:name": "GenericMarker" },
    };
    const parsed = parseOdt(generateOdt({ sections: [], odfExtensions: [extension] }));
    expect(parsed.odfExtensions).toEqual([extension]);
  });

  it("declares OCF subdirectories and packs binary resources", () => {
    const mimeType = "application/vnd.oasis.opendocument.chart";
    const files = {
      "Object 1/content.xml": "<chart:chart/>",
      "Thumbnails/thumbnail.png": new Uint8Array([1, 2, 3]),
    };
    const archive = unzipSync(generateOcf(mimeType, files));
    expect(Array.from(archive["Object 1/content.xml"]!)).toEqual([
      60, 99, 104, 97, 114, 116, 58, 99, 104, 97, 114, 116, 47, 62,
    ]);
    expect(Array.from(archive["Thumbnails/thumbnail.png"]!)).toEqual([1, 2, 3]);
    expect(manifestXml(mimeType, files)).toContain('manifest:full-path="Object 1/"');
    expect(manifestXml(mimeType, files)).toContain('manifest:full-path="Thumbnails/"');
  });

  it("round-trips chart and database subdocuments", () => {
    const chart = {
      title: "Chart",
      chart: {
        class: "chart:bar",
        plotArea: { axes: [], series: [] },
      },
    };
    const parsedChart = parseChartDocument(generateChartDocument(chart));
    expect(parsedChart.title).toBe(chart.title);
    expect(parsedChart.chart).toMatchObject(chart.chart);

    const database = {
      title: "Database",
      body: [{ name: "db:data-source", attributes: { "db:name": "Library" } }],
    };
    const parsedDatabase = parseDatabaseDocument(generateDatabaseDocument(database));
    expect(parsedDatabase).toEqual(database);
  });
});
