import type { PictureOptions } from "@office-open/pptx";
import { describe, expect, it } from "vitest";

import { toDocxPicture, toPptxPicture, toXlsxImage } from "./picture";

describe("toDocxPicture (pptx → docx)", () => {
  it("maps absolute position to offset and passes data through", () => {
    const data = new Uint8Array([1, 2, 3]);
    const pptx: PictureOptions = {
      type: "png",
      data,
      x: 100,
      y: 200,
      width: 300,
      height: 400,
    };
    const docx = toDocxPicture(pptx);
    expect(docx.type).toBe("png");
    expect(docx.data).toBe(data);
    expect(docx.transformation.width).toBe(300);
    expect(docx.transformation.height).toBe(400);
    expect(docx.transformation.offset).toEqual({ left: 100, top: 200 });
  });

  it("omits offset when pptx has no x/y", () => {
    const pptx: PictureOptions = {
      type: "png",
      data: new Uint8Array(),
      width: 10,
      height: 20,
    };
    expect(toDocxPicture(pptx).transformation.offset).toBeUndefined();
  });
});

describe("toXlsxImage (pptx → xlsx)", () => {
  it("maps absolute EMU origin to cell (1,1) and keeps data", () => {
    const data = new Uint8Array([9, 8, 7]);
    const pptx: PictureOptions = { type: "jpg", data, x: 0, y: 0, width: 1, height: 1 };
    const xlsx = toXlsxImage(pptx);
    expect(xlsx.type).toBe("jpg");
    expect(xlsx.data).toBe(data);
    expect(xlsx.col).toBe(1);
    expect(xlsx.row).toBe(1);
  });

  it("drops size (xlsx public input carries none)", () => {
    const xlsx = toXlsxImage({
      type: "png",
      data: new Uint8Array(),
      x: 0,
      y: 0,
      width: 9999,
      height: 9999,
    });
    expect(xlsx).not.toHaveProperty("width");
    expect(xlsx).not.toHaveProperty("height");
  });
});

describe("round-trip pptx → docx → pptx", () => {
  it("preserves binary data, dimensions, and type", () => {
    const data = new Uint8Array([5, 6, 7, 8]);
    const pptx: PictureOptions = {
      type: "png",
      data,
      x: 50,
      y: 60,
      width: 700,
      height: 800,
    };
    const back = toPptxPicture(toDocxPicture(pptx));
    expect(back.data).toBe(data);
    expect(back.width).toBe(700);
    expect(back.height).toBe(800);
    expect(back.type).toBe("png");
    expect(back.x).toBe(50);
    expect(back.y).toBe(60);
  });
});
