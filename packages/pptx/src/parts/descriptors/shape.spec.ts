import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
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
  const el = doc.elements![0];
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
    expect(textBody.text).toBe("Hello");
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
      rotation: 45000,
    });
    expect(result.rotation).toBe(45000);
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
      outline: { color: "000000", width: 12700 },
    });
    const outline = result.outline!;
    expect(outline.color).toBe("000000");
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

  it("round-trips shape with textBody children", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 200,
      height: 100,
      textBody: {
        children: [
          { children: [{ text: "Hello " }, { text: "Bold", bold: true }] },
          { children: [{ text: "World" }] },
        ],
      },
    });
    const textBody = result.textBody!;
    const children = textBody.children! as Record<string, unknown>[];
    expect(children).toHaveLength(2);
    // First paragraph has 2 runs — not simplified to text shorthand
    const para0 = children[0] as { children: Record<string, unknown>[] };
    expect(para0.children).toHaveLength(2);
    expect(para0.children[0].text).toBe("Hello ");
    expect(para0.children[1].text).toBe("Bold");
    expect(para0.children[1].bold).toBe(true);
    // Second paragraph simplified to text shorthand since single run with no properties
    expect(children[1]).toBe("World");
  });

  it("round-trips shape with textBody vertical", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: { text: "Vertical", vertical: "vert" },
    });
    const textBody = result.textBody!;
    expect(textBody.vertical).toBe("vert");
  });

  it("round-trips shape with textBody anchor", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: { text: "Centered", anchor: "center" },
    });
    const textBody = result.textBody!;
    expect(textBody.anchor).toBe("center");
  });

  it("round-trips shape with textBody autoFit", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: { text: "Auto", autoFit: "normal" },
    });
    const textBody = result.textBody!;
    expect(textBody.autoFit).toBe("normal");
  });

  it("round-trips shape with textBody margins", () => {
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      textBody: {
        text: "Margins",
        margins: { top: 1000, bottom: 2000, left: 3000, right: 4000 },
      },
    });
    const textBody = result.textBody!;
    const margins = textBody.margins!;
    expect(margins.top).toBe(1000);
    expect(margins.bottom).toBe(2000);
    expect(margins.left).toBe(3000);
    expect(margins.right).toBe(4000);
  });

  it("round-trips shape reflection effect with all fields", () => {
    // CT_ReflectionEffect has 14 attrs; parse must read all (was dropping
    // stPos/endPos/fadeDir/sx/sy/kx/ky/algn/rotWithShape) and invert the
    // unit scaling applied by toReflectionCore.
    const result = roundTrip({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      effects: {
        reflection: {
          blurRadius: 12700,
          distance: 50000,
          direction: 5400000,
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
          rotateWithShape: false,
        },
      },
    });
    const reflection = result.effects!.reflection!;
    expect(reflection.blurRadius).toBe(12700);
    expect(reflection.distance).toBe(50000);
    expect(reflection.direction).toBe(5400000);
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
    expect(reflection.rotateWithShape).toBe(false);
  });
});
