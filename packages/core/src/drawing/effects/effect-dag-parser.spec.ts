import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { ReadContext, WriteContext } from "../../descriptor";
import { createEffectDag } from "./effect-dag";
import { parseEffectContainer } from "./effect-dag-parser";

const SOURCE =
  '<a:effectDag xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'type="tree" name="fx">' +
  '<a:alphaBiLevel thresh="50%"/>' +
  '<a:blend blend="over"><a:cont><a:glow rad="40000">' +
  '<a:srgbClr val="FF0000"/></a:glow></a:cont></a:blend>' +
  '<a:clrChange useA="0"><a:clrFrom><a:srgbClr val="000000"/></a:clrFrom>' +
  '<a:clrTo><a:schemeClr val="accent1"/></a:clrTo></a:clrChange>' +
  '<a:fill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fill>' +
  '<a:fillOverlay blend="screen"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:fillOverlay>' +
  '<a:relOff tx="10%" ty="-5%"/>' +
  "</a:effectDag>";

const BLIP_SOURCE =
  '<a:effectDag xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<a:fill><a:blipFill><a:blip r:embed="rId1"/></a:blipFill></a:fill>' +
  '<a:fillOverlay blend="screen"><a:blipFill><a:blip r:embed="rId2"/></a:blipFill></a:fillOverlay>' +
  "</a:effectDag>";

function parseSource(xml: string, ctx: ReadContext = {} as ReadContext) {
  const element = parseXml(xml).elements?.[0];
  if (!element) throw new Error("expected effect DAG root");
  return parseEffectContainer(element, ctx);
}

describe("effect DAG parser", () => {
  const relationshipContext = {
    resolveRelationship: () => undefined,
    getPart: () => undefined,
    getRaw: () => undefined,
  } satisfies ReadContext;

  it("parses all modeled effect families into typed options", () => {
    const parsed = parseSource(SOURCE, relationshipContext);
    expect(parsed.type).toBe("tree");
    expect(parsed.name).toBe("fx");
    expect(parsed.alphaBiLevel).toEqual({ threshold: 50 });
    expect(parsed.blend?.blend).toBe("over");
    expect(parsed.fill?.fill).toMatchObject({ type: "solid", color: { value: "FF0000" } });
    expect(parsed.fillOverlay?.fill).toMatchObject({ type: "solid", color: { value: "00FF00" } });
    expect(parsed.blend?.container.glow?.radius).toBe(40000);
    expect(parsed.colorChange?.useAlpha).toBe(false);
    expect(parsed.relativeOffset).toEqual({ translateX: 10, translateY: -5 });
  });

  it("round-trips generated effect DAGs without raw XML", () => {
    const parsed = parseSource(SOURCE, relationshipContext);
    const generated = createEffectDag(parsed);
    expect(parseSource(generated, relationshipContext)).toEqual(parsed);
    expect(generated).toContain('thresh="50000"');
  });

  it("registers effect blip-fill media through the active write context", () => {
    const added: string[] = [];
    const context: WriteContext = {
      addRelationship: () => "rId1",
      addMedia: (_data, _type, fileName) => {
        added.push(fileName ?? "");
        return `{${fileName}}`;
      },
      addHyperlink: () => {},
    };
    const generated = createEffectDag(
      {
        fill: {
          fill: {
            type: "blip",
            data: new Uint8Array([1, 2, 3]),
            imageType: "png",
            fileName: "fill.png",
          },
        },
      },
      context,
    );
    expect(added).toEqual(["fill.png"]);
    expect(generated).toContain('r:embed="{fill.png}"');
  });

  it("resolves effect blip fills through the owning part relationships", () => {
    const parsed = parseSource(BLIP_SOURCE, {
      resolveRelationship: (rId: string) =>
        rId === "rId1" ? "media/fill.png" : rId === "rId2" ? "media/overlay.png" : undefined,
      getRaw: (path: string) => (path.endsWith(".png") ? new Uint8Array([1, 2, 3]) : undefined),
    } as ReadContext);
    expect(parsed.fill?.fill).toMatchObject({ type: "blip", fileName: "fill.png" });
    expect(parsed.fillOverlay?.fill).toMatchObject({ type: "blip", fileName: "overlay.png" });
  });
});
