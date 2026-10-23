import { parseXmlElement } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentation, parsePresentation } from "./index";
import type { OdpPresentationOptions } from "./options";
import {
  parseFrameNode,
  parseMeasureNode,
  parsePageThumbnailNode,
  parseSceneNode,
  slideNodeXml,
  type OdpSlideNodePayload,
} from "./slide-nodes";

function roundTrip(child: OdpSlideNodePayload): OdpSlideNodePayload {
  const first = slideNodeXml(child, [], []);
  const parsed = parseNode(first);
  expect(slideNodeXml(parsed, [], [])).toBe(first);
  return parsed;
}

function parseNode(xml: string): OdpSlideNodePayload {
  const element = parseXmlElement(xml);
  if (!element) throw new TypeError("Expected an ODP slide node");
  if (element.name === "draw:frame") return parseFrameNode(element);
  if (element.name === "draw:measure") return parseMeasureNode(element);
  if (element.name === "draw:page-thumbnail") return parsePageThumbnailNode(element);
  if (element.name === "dr3d:scene") return parseSceneNode(element);
  throw new TypeError(`Unexpected ODP slide node: ${element.name}`);
}

describe("ODP typed slide nodes", () => {
  it("round-trips a mixed frame with RNG frame children", () => {
    const child = parseNode(
      '<draw:frame draw:name="Frame" draw:z-index="3" draw:caption-id="caption" ' +
        'svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm" style:rel-width="scale">' +
        '<draw:applet draw:code="Clock" draw:may-script="true" xlink:type="simple" ' +
        'xlink:href="applets/clock" xlink:show="embed" xlink:actuate="onLoad">' +
        '<draw:param draw:name="rate" draw:value="1"/></draw:applet>' +
        '<draw:floating-frame draw:frame-name="Browser" xlink:type="simple" xlink:href="https://example.com/"/>' +
        '<draw:plugin draw:mime-type="application/x-example" xlink:type="simple" xlink:href="plugins/example">' +
        '<draw:param draw:name="mode" draw:value="compact"/></draw:plugin>' +
        '<draw:text-box fo:min-height="1cm" draw:auto-grow-height="true"><text:p>Frame text</text:p></draw:text-box>' +
        "<table:table><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table>" +
        '<office:event-listeners><script:event-listener script:event-name="click" script:language="JavaScript" ' +
        'script:macro-name="Run"/></office:event-listeners>' +
        "<svg:title>Accessible frame</svg:title><svg:desc>Frame description</svg:desc>" +
        "</draw:frame>",
    );
    expect(child).toEqual(
      roundTrip({
        frame: {
          name: "Frame",
          zIndex: 3,
          captionId: "caption",
          x: 360000,
          y: 720000,
          width: 1440000,
          height: 1080000,
          relativeWidth: "scale",
          title: "Accessible frame",
          description: "Frame description",
          listeners: [
            { target: "macro", eventName: "click", language: "JavaScript", macroName: "Run" },
          ],
          contents: [
            {
              kind: "applet",
              applet: {
                code: "Clock",
                mayScript: true,
                data: { href: "applets/clock", show: "embed", actuate: "onLoad" },
                parameters: [{ name: "rate", value: "1" }],
              },
            },
            {
              kind: "floating-frame",
              floatingFrame: {
                frameName: "Browser",
                data: { href: "https://example.com/" },
              },
            },
            {
              kind: "plugin",
              plugin: {
                mimeType: "application/x-example",
                data: { href: "plugins/example" },
                parameters: [{ name: "mode", value: "compact" }],
              },
            },
            {
              kind: "text-box",
              textBox: {
                minHeight: 360000,
                autoGrowHeight: true,
                text: { paragraphs: ["Frame text"] },
              },
            },
            {
              kind: "table",
              table: { columnWidths: [], rows: [{ cells: [{ children: ["Cell"] }] }] },
            },
          ],
        },
      }),
    );
  });

  it("round-trips all draw:measure required attributes and text", () => {
    const child = parseNode(
      '<draw:measure draw:id="measurement" draw:layer="layout" draw:transform="matrix(1 0 0 1 0 0)" ' +
        'svg:x1="1cm" svg:y1="2cm" svg:x2="4cm" svg:y2="2cm" draw:measure-align="left" ' +
        'draw:measure-vertical-align="above"><draw:glue-point draw:id="1" svg:x="50%" svg:y="0" ' +
        'draw:escape-direction="top"/><svg:title>Length</svg:title><text:p>12 cm</text:p></draw:measure>',
    );
    expect(child).toEqual(
      roundTrip({
        measure: {
          drawId: "measurement",
          layer: "layout",
          transform: "matrix(1 0 0 1 0 0)",
          x1: 360000,
          y1: 720000,
          x2: 1440000,
          y2: 720000,
          measureAlign: "left",
          measureVerticalAlign: "above",
          title: "Length",
          gluePoints: [{ id: 1, x: "50%", y: 0, escapeDirection: "top" }],
          text: { paragraphs: ["12 cm"] },
        },
      }),
    );
  });

  it("round-trips a page thumbnail and its optional page reference", () => {
    const child = parseNode(
      '<draw:page-thumbnail draw:name="Thumbnail" presentation:class="notes" draw:page-number="2" ' +
        'svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="1.5cm"><svg:desc>Second page</svg:desc></draw:page-thumbnail>',
    );
    expect(child).toEqual(
      roundTrip({
        pageThumbnail: {
          name: "Thumbnail",
          presentationClass: "notes",
          pageNumber: 2,
          x: 360000,
          y: 360000,
          width: 720000,
          height: 540000,
          description: "Second page",
        },
      }),
    );
  });

  it("round-trips scene camera state, lights, nested 3D objects, and glue points", () => {
    const child = parseNode(
      '<dr3d:scene draw:name="Scene" dr3d:transform="matrix(1 0 0 1 0 0)" dr3d:vrp="0cm 0cm 10cm" ' +
        'dr3d:vpn="0cm 0cm -1cm" dr3d:vup="0cm 1cm 0cm" dr3d:projection="perspective" ' +
        'dr3d:distance="10cm" dr3d:shadow-slant="15" dr3d:shade-mode="phong" dr3d:ambient-color="#101010" ' +
        'svg:x="1cm" svg:y="1cm" svg:width="3cm" svg:height="3cm">' +
        '<dr3d:light dr3d:direction="1cm 1cm 1cm" dr3d:diffuse-color="#ffffff" dr3d:enabled="true" dr3d:specular="false"/>' +
        '<dr3d:cube dr3d:min-edge="0cm 0cm 0cm" dr3d:max-edge="1cm 1cm 1cm"/>' +
        '<dr3d:sphere dr3d:center="2cm 2cm 2cm" dr3d:size="1cm 1cm 1cm"/>' +
        '<dr3d:extrude svg:d="M 0 0 L 1 1" svg:viewBox="0 0 1 1" dr3d:depth="1cm" dr3d:close-front="true"/>' +
        '<dr3d:rotate svg:d="M 0 0 L 1 1" svg:viewBox="0 0 1 1" dr3d:horizontal="true"/>' +
        '<dr3d:scene dr3d:projection="parallel"/>' +
        '<draw:glue-point draw:id="2" svg:x="1cm" svg:y="1cm" draw:escape-direction="left"/>' +
        "</dr3d:scene>",
    );
    expect(child).toEqual(
      roundTrip({
        scene: {
          name: "Scene",
          transform: "matrix(1 0 0 1 0 0)",
          viewRecommendationPoint: { x: 0, y: 0, z: 3600000 },
          viewPlaneNormal: { x: 0, y: 0, z: -360000 },
          viewUp: { x: 0, y: 360000, z: 0 },
          projection: "perspective",
          distance: 3600000,
          shadowSlant: 15,
          shadeMode: "phong",
          ambientColor: "#101010",
          x: 360000,
          y: 360000,
          width: 1080000,
          height: 1080000,
          lights: [
            {
              direction: { x: 360000, y: 360000, z: 360000 },
              diffuseColor: "#ffffff",
              enabled: true,
              specular: false,
            },
          ],
          objects: [
            {
              type: "cube",
              minEdge: { x: 0, y: 0, z: 0 },
              maxEdge: { x: 360000, y: 360000, z: 360000 },
            },
            {
              type: "sphere",
              center: { x: 720000, y: 720000, z: 720000 },
              size: { x: 360000, y: 360000, z: 360000 },
            },
            {
              type: "extrude",
              path: "M 0 0 L 1 1",
              viewBox: [0, 0, 1, 1],
              depth: 360000,
              closeFront: true,
            },
            { type: "rotate", path: "M 0 0 L 1 1", viewBox: [0, 0, 1, 1], horizontal: true },
            { projection: "parallel" },
          ],
          gluePoints: [{ id: 2, x: 360000, y: 360000, escapeDirection: "left" }],
        },
      }),
    );
  });

  it("rejects unmodeled frame children instead of preserving raw XML", () => {
    const element = parseXmlElement("<draw:frame><draw:unmodeled/></draw:frame>");
    if (!element) throw new TypeError("Expected a frame");
    expect(() => parseFrameNode(element)).toThrow("Unknown child element");
  });

  it("preserves typed ODF-only nodes beside canonical slide children", async () => {
    const scene = { scene: { projection: "parallel" as const } };
    const options: OdpPresentationOptions = {
      slides: [
        {
          children: [{ line: { x1: 0, y1: 0, x2: 360000, y2: 360000 } }],
          odpNodes: [{ position: 0, node: scene }],
        },
      ],
    };
    const parsed = await parsePresentation(await generatePresentation(options));
    const slide = options.slides?.[0];
    if (!slide) throw new TypeError("Expected an ODP slide");
    expect(parsed.slides?.[0]?.children).toEqual(slide.children);
    expect(parsed.slides?.[0]?.odpNodes).toEqual(slide.odpNodes);
    expect(await parsePresentation(await generatePresentation(parsed))).toEqual(parsed);
  });
});
