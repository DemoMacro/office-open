import type { CustomGeometryOptions } from "@office-open/core/drawing";
import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { OdfSchemaError } from "../errors";
import { odfEnhancedGeometryXml, parseOdfEnhancedGeometry } from "./enhanced-geometry";
import { parseOdfPrimitiveGeometry } from "./primitive-geometry";

function parseGeometry(xml: string): CustomGeometryOptions | undefined {
  const element = parse(xml, { ignoreDeclaration: true }).elements?.find(
    (item) => item.type === "element",
  );
  return parseOdfEnhancedGeometry(element);
}

describe("ODF enhanced geometry", () => {
  it("projects primitive paths and point lists into canonical geometry", () => {
    const path = parse(
      '<draw:path svg:viewBox="0 0 100 50" svg:d="M 0 0 C 25 0 75 50 100 50 Z"/>',
      { ignoreDeclaration: true },
    ).elements?.find((item) => item.type === "element");
    const polygon = parse('<draw:polygon svg:viewBox="0 0 10 10" draw:points="0,0 10,0 5,10"/>', {
      ignoreDeclaration: true,
    }).elements?.find((item) => item.type === "element");
    const polyline = parse('<draw:polyline svg:viewBox="0 0 10 10" svg:points="0 0 10 10"/>', {
      ignoreDeclaration: true,
    }).elements?.find((item) => item.type === "element");

    expect(parseOdfPrimitiveGeometry(path!, 1000, 500)).toEqual({
      pathList: [
        {
          w: 1000,
          h: 500,
          commands: [
            { command: "moveTo", point: { x: "0", y: "0" } },
            {
              command: "cubicBezTo",
              points: [
                { x: "250", y: "0" },
                { x: "750", y: "500" },
                { x: "1000", y: "500" },
              ],
            },
            { command: "close" },
          ],
        },
      ],
    });
    expect(parseOdfPrimitiveGeometry(polygon!, 1000, 1000)?.pathList?.[0]?.commands).toEqual([
      { command: "moveTo", point: { x: "0", y: "0" } },
      { command: "lineTo", point: { x: "1000", y: "0" } },
      { command: "lineTo", point: { x: "500", y: "1000" } },
      { command: "close" },
    ]);
    expect(parseOdfPrimitiveGeometry(polyline!, 1000, 1000)?.pathList?.[0]?.commands).toEqual([
      { command: "moveTo", point: { x: "0", y: "0" } },
      { command: "lineTo", point: { x: "1000", y: "1000" } },
    ]);
  });

  it("round-trips guides, handles, connection sites, text rectangles, and multiple paths", () => {
    const geometry: CustomGeometryOptions = {
      adjustmentValues: [{ name: "modifier0", formula: "val 25" }],
      guides: [{ name: "half", formula: "*/ $0 1 2" }],
      adjustHandles: [
        {
          type: "xy",
          minX: "0",
          maxX: "100000",
          minY: "0",
          maxY: "100000",
          position: { x: "$0", y: "half" },
        },
        {
          type: "polar",
          guideRefRadius: "radius",
          guideRefAngle: "angle",
          minRadius: "0",
          maxRadius: "100000",
          position: { x: "50000", y: "50000" },
        },
      ],
      connectionSites: [{ angle: "0", position: { x: "0", y: "50000" } }],
      textRectangle: { left: "0", top: "0", right: "100000", bottom: "100000" },
      pathList: [
        {
          w: 100000,
          h: 100000,
          commands: [
            { command: "moveTo", point: { x: "0", y: "0" } },
            { command: "lineTo", point: { x: "modifier0", y: "100000" } },
            { command: "close" },
          ],
        },
        {
          w: 100000,
          h: 100000,
          commands: [
            { command: "moveTo", point: { x: "0", y: "0" } },
            {
              command: "quadBezTo",
              points: [
                { x: "50000", y: "0" },
                { x: "100000", y: "100000" },
              ],
            },
          ],
        },
      ],
    };

    expect(parseGeometry(odfEnhancedGeometryXml(geometry))).toEqual(geometry);
  });

  it("rejects ODF geometry attributes without a canonical field", () => {
    expect(() =>
      parseGeometry(
        '<draw:enhanced-geometry xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:glue-point-type="rectangle"/>',
      ),
    ).toThrow(OdfSchemaError);
    expect(() =>
      parseGeometry(
        '<draw:enhanced-geometry xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" svg:viewBox="1 0 100 100"/>',
      ),
    ).toThrow(OdfSchemaError);
  });

  it("rejects canonical geometry fields that ODF cannot represent", () => {
    expect(() =>
      odfEnhancedGeometryXml({
        connectionSites: [{ angle: "90", position: { x: "0", y: "0" } }],
      }),
    ).toThrow(OdfSchemaError);
    expect(() =>
      odfEnhancedGeometryXml({
        adjustHandles: [
          {
            type: "xy",
            guideRefX: "x",
            position: { x: "0", y: "0" },
          },
        ],
      }),
    ).toThrow(OdfSchemaError);
  });
});
