import { parse } from "@office-open/xml";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import {
  generateOcf,
  manifestXml,
  ODF_ELEMENT_NAMES,
  parseOdfNodes,
  readOcf,
  serializeOdfNodes,
  type OdfXmlNode,
} from "./index";

describe("OCF runtime", () => {
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
    expect(parseOdfNodes(document.elements?.[0])).toEqual(nodes);
  });

  it("declares OCF subdirectories and packs binary resources", () => {
    const mimeType = "application/vnd.oasis.opendocument.chart";
    const files = {
      "Object 1/content.xml": "<chart:chart/>",
      "Thumbnails/thumbnail.png": new Uint8Array([1, 2, 3]),
    };
    const archive = generateOcf(mimeType, files);
    expect(archive.subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    expect(manifestXml(mimeType, files)).toContain('manifest:full-path="Object 1/"');
    expect(manifestXml(mimeType, files)).toContain('manifest:full-path="Thumbnails/"');
  });

  it("opens a manifest-identified package without a mimetype entry", () => {
    const archive = zipSync({
      "content.xml": strToU8(
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      ),
      "META-INF/manifest.xml": strToU8(
        '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
      ),
    });
    expect(() => readOcf(archive, "application/vnd.oasis.opendocument.text")).not.toThrow();
  });
});
