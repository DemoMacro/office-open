import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { ReflectionEffectOptions } from "@office-open/core/drawingml";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it, beforeEach } from "vite-plus/test";

import { shapeDesc, resetShapeIdCounter } from "./shape";
import type { ShapeDescriptorOptions } from "./shape";

// ── Mock PPTX write context ──

class MockWriteContext {
  private _nextRelId = 1;
  addRelationship() {
    return `rId${this._nextRelId++}`;
  }
  addMedia() {
    return "";
  }
  addHyperlink() {}
  addImage() {}
}

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: ShapeDescriptorOptions) {
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
      fill: { type: "solid", color: "4472C4" },
    });
    const fill = result.fill! as { type: string; color: { value: string } };
    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("4472C4");
  });

  it("round-trips shape with geometry", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      geometry: "ellipse",
    });
    expect(result.geometry).toBe("ellipse");
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

  it("round-trips shape with blackWhiteMode", () => {
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
      outline: { type: "solidFill", color: { value: "000000" }, width: 12700 },
    });
    const outline = result.outline as {
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
        fontReference: { index: 0, color: "333333" },
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
      textBody: { text: "Vertical", bodyProperties: { vert: "vert" } },
    });
    const textBody = result.textBody!;
    expect(textBody.bodyProperties?.vert).toBe("vert");
  });

  it("round-trips shape with textBody anchor", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: { text: "Centered", bodyProperties: { anchor: "ctr" } },
    });
    const textBody = result.textBody!;
    expect(textBody.bodyProperties?.anchor).toBe("ctr");
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
    });
    const reflection = result.effects!.reflection as ReflectionEffectOptions;
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
