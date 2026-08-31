import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { ReflectionEffectOptions } from "@office-open/core/drawing";
import { parse as parseXml } from "@office-open/xml";
import type { ShapeOptions } from "@shared/shape/shape";
import { describe, expect, it, beforeEach } from "vite-plus/test";

import { shapeDesc, pictureDesc, resetShapeIdCounter } from "./shape";

// ── Mock PPTX write context ──

class MockWriteContext {
  registerShapeId() {}
  private _nextRelId = 1;
  addRelationship() {
    return `rId${this._nextRelId++}`;
  }
  addMedia() {
    return "";
  }
  addHyperlink(_key: string, target: { url?: string; slide?: number; tooltip?: string }) {
    this.lastHyperlink = target;
  }
  addImage() {}
  lastHyperlink?: { url?: string; slide?: number; tooltip?: string };
}

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: ShapeOptions) {
  const writeCtx = new MockWriteContext() as unknown as WriteContext;
  const xml = shapeDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return shapeDesc.parse(el, readCtx);
}

describe("shapeDesc round-trip", () => {
  beforeEach(() => {
    resetShapeIdCounter(2);
  });

  it("round-trips basic shape with position", () => {
    const result = roundTrip({ x: 100, y: 200, width: 400, height: 300 });
    expect(result.x).toBe(100);
    expect(result.y).toBe(200);
    expect(result.width).toBe(400);
    expect(result.height).toBe(300);
  });

  it("round-trips shape with id and name", () => {
    const result = roundTrip({ id: 42, name: "MyShape", x: 0, y: 0, width: 100, height: 100 });
    expect(result.id).toBe(42);
    expect(result.name).toBe("MyShape");
  });

  it("round-trips shape with textBody text", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 200,
      height: 100,
      textBody: { text: "Hello" },
    });
    const textBody = result.textBody!;
    expect((textBody.paragraphs?.[0] as { text?: string })?.text).toBe("Hello");
  });

  it("omits p:txBody for a textBody-less shape (e.g. sldImg placeholder)", () => {
    // txBody is optional in CT_Shape. A shape without textBody — like the notes
    // slide-image placeholder — must round-trip without a spurious empty body.
    const writeCtx = new MockWriteContext() as unknown as WriteContext;
    const xml = shapeDesc.stringify(
      { id: 2, name: "Picture", x: 0, y: 0, width: 100, height: 100 },
      writeCtx,
    )!;
    expect(xml).not.toContain("<p:txBody");

    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = shapeDesc.parse(el, readCtx);
    expect(result.textBody).toBeUndefined();
  });

  it("round-trips shape with solidFill", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      properties: { fill: { type: "solid", color: "4472C4" } },
    });
    const fill = result.properties!.fill! as { type: string; color: { value: string } };
    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("4472C4");
  });

  it("round-trips shape with geometry", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      properties: { geometry: "ellipse" },
    });
    expect(result.properties?.geometry).toBe("ellipse");
  });

  it("round-trips shape with placeholder", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      placeholder: "title",
    });
    expect(result.placeholder).toBe("title");
  });

  it("round-trips shape with placeholderIndex", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      placeholder: "body",
      placeholderIndex: 1,
    });
    expect(result.placeholder).toBe("body");
    expect(result.placeholderIndex).toBe(1);
  });

  it("round-trips shape with rotation", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      rotation: 45,
    });
    expect(result.rotation).toBe(45);
  });

  it("round-trips shape with flipHorizontal", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      flipHorizontal: true,
    });
    expect(result.flipHorizontal).toBe(true);
  });

  it("round-trips shape with flipVertical", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      flipVertical: true,
    });
    expect(result.flipVertical).toBe(true);
  });

  it("round-trips shape click hyperlink on cNvPr", () => {
    const writeCtx = new MockWriteContext();
    const xml = shapeDesc.stringify(
      {
        x: 0,
        y: 0,
        width: 100,
        height: 100,
        hyperlink: { url: "https://example.com", tooltip: "Go" },
      },
      writeCtx as unknown as WriteContext,
    )!;
    // Registered on the write context for the compiler to emit the relationship
    expect(writeCtx.lastHyperlink?.url).toBe("https://example.com");
    expect(xml).toContain("<a:hlinkClick ");

    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = shapeDesc.parse(el, readCtx);
    // The mock read context resolves no relationships, so the target comes back
    // as the placeholder referenceId; tooltip round-trips directly.
    expect(result.hyperlink?.referenceId).toBeDefined();
    expect(result.hyperlink?.tooltip).toBe("Go");
  });

  it("round-trips shape with blackWhiteMode on spPr", () => {
    // @bwMode is a CT_ShapeProperties attribute; CT_Shape itself only has
    // @useBgFill.
    const xml = shapeDesc.stringify(
      { x: 0, y: 0, width: 100, height: 100, blackWhiteMode: "gray" },
      { registerShapeId() {} } as never,
    );
    expect(xml).toContain('<p:spPr bwMode="gray">');
    expect(xml).not.toMatch(/<p:sp [^>]*bwMode/);

    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      blackWhiteMode: "gray",
    });
    expect(result.blackWhiteMode).toBe("gray");
  });

  it("round-trips shape with useBackgroundFill", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      useBackgroundFill: true,
    });
    expect(result.useBackgroundFill).toBe(true);
  });

  it("round-trips shape with outline", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      properties: { outline: { type: "solidFill", color: { value: "000000" }, width: 12700 } },
    });
    const outline = result.properties!.outline as {
      type: string;
      color: { value: string };
      width: number;
    };
    expect(outline.type).toBe("solidFill");
    expect(outline.color.value).toBe("000000");
    expect(outline.width).toBe(12700);
  });

  it("round-trips shape with style", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      style: {
        lineReference: { index: 2, color: "333333" },
        fillReference: { index: 1 },
        effectReference: { index: 0 },
        fontReference: { collection: "minor", color: "333333" },
      },
    });
    const style = result.style!;
    const lineRef = style.lineReference!;
    expect(lineRef.index).toBe(2);
    expect(lineRef.color).toBe("333333");
    const fillRef = style.fillReference!;
    expect(fillRef.index).toBe(1);
  });

  it("round-trips shape with locking", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      locking: { noGrp: true, noSelect: true, noRot: true },
    });
    const locking = result.locking!;
    expect(locking.noGrp).toBe(true);
    expect(locking.noSelect).toBe(true);
    expect(locking.noRot).toBe(true);
  });

  it("round-trips shape with isPhoto", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      isPhoto: true,
    });
    expect(result.isPhoto).toBe(true);
  });

  it("round-trips shape with hasCustomPrompt", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      placeholder: "body",
      hasCustomPrompt: true,
    });
    expect(result.hasCustomPrompt).toBe(true);
  });

  it("round-trips shape with textBody paragraphs", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 200,
      height: 100,
      textBody: {
        paragraphs: [
          { children: [{ text: "Hello " }, { text: "Bold", bold: true }] },
          { children: [{ text: "World" }] },
        ],
      },
    });
    const textBody = result.textBody!;
    const paragraphs = textBody.paragraphs!;
    expect(paragraphs).toHaveLength(2);
    // First paragraph has 2 runs — not simplified to text shorthand
    const para0 = paragraphs[0] as { children: Record<string, unknown>[] };
    expect(para0.children).toHaveLength(2);
    const [run0, run1] = para0.children;
    expect(run0?.text).toBe("Hello ");
    expect(run1?.text).toBe("Bold");
    expect(run1?.bold).toBe(true);
    // Second paragraph: single run, no properties → text shorthand
    const para1 = paragraphs[1] as { text?: string };
    expect(para1.text).toBe("World");
  });

  it("round-trips shape with textBody vertical", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: { text: "Vertical", bodyProperties: { vertical: "vertical" } },
    });
    const textBody = result.textBody!;
    expect(textBody.bodyProperties?.vertical).toBe("vertical");
  });

  it("round-trips shape with textBody anchor", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: { text: "Centered", bodyProperties: { anchor: "center" } },
    });
    const textBody = result.textBody!;
    expect(textBody.bodyProperties?.anchor).toBe("center");
  });

  it("round-trips shape with textBody autofit", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: { text: "Auto", bodyProperties: { normAutofit: {} } },
    });
    const textBody = result.textBody!;
    expect(textBody.bodyProperties?.normAutofit).toEqual({});
  });

  it("round-trips shape with textBody margins", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: {
        text: "Margins",
        bodyProperties: { margins: { top: 1000, bottom: 2000, left: 3000, right: 4000 } },
      },
    });
    const textBody = result.textBody!;
    const bodyProperties = textBody.bodyProperties!;
    expect(bodyProperties.tIns).toBe(1000);
    expect(bodyProperties.bIns).toBe(2000);
    expect(bodyProperties.lIns).toBe(3000);
    expect(bodyProperties.rIns).toBe(4000);
  });

  it("round-trips shape reflection effect with all fields", () => {
    // CT_ReflectionEffect has 14 attrs; core reflectionDesc reads all and
    // stores XSD-native values (percentages *1000, angles *60000).
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      properties: {
        effects: {
          reflection: {
            blurRadius: 12700,
            distance: 50000,
            direction: 90,
            startAlpha: 60,
            startPosition: 10,
            endAlpha: 0,
            endPosition: 90,
            fadeDirection: 90,
            scaleX: 50,
            scaleY: 75,
            skewX: 45,
            skewY: 30,
            alignment: "bottomLeft",
            rotWithShape: false,
          },
        },
      },
    });
    const reflection = result.properties!.effects!.reflection as ReflectionEffectOptions;
    expect(reflection.blurRadius).toBe(12700);
    expect(reflection.distance).toBe(50000);
    expect(reflection.direction).toBe(90);
    expect(reflection.startAlpha).toBe(60);
    expect(reflection.startPosition).toBe(10);
    expect(reflection.endAlpha).toBe(0);
    expect(reflection.endPosition).toBe(90);
    expect(reflection.fadeDirection).toBe(90);
    expect(reflection.scaleX).toBe(50);
    expect(reflection.scaleY).toBe(75);
    expect(reflection.skewX).toBe(45);
    expect(reflection.skewY).toBe(30);
    expect(reflection.alignment).toBe("bottomLeft");
    expect(reflection.rotWithShape).toBe(false);
  });
});

describe("pictureDesc round-trip", () => {
  // pictureDesc.stringify registers the image via addImage and reads the
  // canonical fileName back — mirror the entry instead of returning "".
  const picWriteCtx = {
    registerShapeId() {},
    addRelationship: () => "rId1",
    addMedia: () => "",
    addImage: (_key: string, entry: { fileName: string }) => entry,
  } as unknown as WriteContext;

  function roundTripPicture(opts: Parameters<typeof pictureDesc.stringify>[0]) {
    const xml = pictureDesc.stringify(opts, picWriteCtx)!;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    return pictureDesc.parse(el, readCtx);
  }

  it("round-trips picture flip and rotation from a:xfrm", () => {
    const result = roundTripPicture({
      id: 7,
      name: "Flipped",
      data: "dummy",
      type: "png",
      x: 10,
      y: 20,
      width: 100,
      height: 50,
      flipHorizontal: true,
      flipVertical: true,
      rotation: 90,
    });
    expect(result.flipHorizontal).toBe(true);
    expect(result.flipVertical).toBe(true);
    expect(result.rotation).toBe(90);
  });

  it("emits a linked-only blip (r:link, no media registration)", () => {
    const imageLinks: string[] = [];
    const ctx = {
      registerShapeId() {},
      addRelationship: () => "rId1",
      addMedia: () => "",
      addImage: () => {
        throw new Error("linked-only picture must not register media");
      },
      addImageLink: (url: string) => {
        imageLinks.push(url);
        return `img-link_${imageLinks.length}`;
      },
    } as unknown as WriteContext;
    const xml = pictureDesc.stringify(
      {
        id: 3,
        name: "Linked",
        type: "gif",
        sourceUrl: "http://example.com/logo.gif",
        x: 0,
        y: 0,
        width: 100,
        height: 100,
      },
      ctx,
    )!;
    expect(xml).toContain('r:link="{img-link:img-link_1}"');
    expect(xml).not.toContain("r:embed");
    expect(imageLinks).toEqual(["http://example.com/logo.gif"]);
  });

  it("parses a linked-only blip without fabricating bytes", () => {
    const xml =
      '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:nvPicPr><p:cNvPr id="2" name="Rectangle 3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:link="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      "<p:spPr/></p:pic>";
    const linkedReadCtx = {
      resolveRelationship: (rId: string) =>
        rId === "rId2" ? "http://www.google.com/intl/en/images/logo.gif" : undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
    } as unknown as ReadContext;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = pictureDesc.parse(el, linkedReadCtx);
    expect(result.sourceUrl).toBe("http://www.google.com/intl/en/images/logo.gif");
    expect(result.data).toBeUndefined();
    expect(result.type).toBe("gif");
  });

  it("round-trips blip compression and omits cstate when unset", () => {
    const withCompression = pictureDesc.stringify(
      {
        id: 5,
        name: "C",
        data: "dummy",
        type: "png",
        compression: "print",
        x: 0,
        y: 0,
        width: 10,
        height: 10,
      },
      picWriteCtx,
    )!;
    expect(withCompression).toContain('cstate="print"');
    const parsed = pictureDesc.parse(parseXml(withCompression).elements![0]!, readCtx);
    expect(parsed.compression).toBe("print");

    const without = pictureDesc.stringify(
      { id: 6, name: "N", data: "dummy", type: "png", x: 0, y: 0, width: 10, height: 10 },
      picWriteCtx,
    )!;
    expect(without).not.toContain("cstate");
  });
});
