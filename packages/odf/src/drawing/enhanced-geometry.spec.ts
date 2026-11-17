import type { CustomGeometryOptions } from "@office-open/core/drawing";
import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { OdfSchemaError } from "../errors";
import { odfEnhancedGeometryXml, parseOdfEnhancedGeometry } from "./enhanced-geometry";

function parseGeometry(xml: string): CustomGeometryOptions | undefined {
  const element = parse(xml, { ignoreDeclaration: true }).elements?.find(
    (item) => item.type === "element",
  );
  return parseOdfEnhancedGeometry(element);
}

describe("ODF enhanced geometry", () => {
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
