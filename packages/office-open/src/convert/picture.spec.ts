import type { PictureOptions as DocxPictureOptions } from "@office-open/docx";
import type { PictureOptions as PptxPictureOptions } from "@office-open/pptx";
import type { PictureOptions as XlsxPictureOptions } from "@office-open/xlsx";
import { describe, expect, it } from "vitest";

import { toDocxPicture, toPptxPicture, toXlsxPicture } from "./picture";

describe("toDocxPicture (pptx → docx)", () => {
  it("maps absolute position to offset and passes data through", () => {
    const data = new Uint8Array([1, 2, 3]);
    const pptx: PptxPictureOptions = {
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
    const pptx: PptxPictureOptions = {
      type: "png",
      data: new Uint8Array(),
      width: 10,
      height: 20,
    };
    expect(toDocxPicture(pptx).transformation.offset).toBeUndefined();
  });
});

describe("toXlsxPicture (pptx → xlsx)", () => {
  it("maps absolute EMU origin to cell (1,1) and keeps data", () => {
    const data = new Uint8Array([9, 8, 7]);
    const pptx: PptxPictureOptions = { type: "jpg", data, x: 0, y: 0, width: 1, height: 1 };
    const xlsx = toXlsxPicture(pptx);
    expect(xlsx.type).toBe("jpg");
    expect(xlsx.data).toBe(data);
    expect(xlsx.col).toBe(1);
    expect(xlsx.row).toBe(1);
  });

  it("drops size (xlsx public input carries none)", () => {
    const xlsx = toXlsxPicture({
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
    const pptx: PptxPictureOptions = {
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

describe("cross-format cNvPr (alt text) preservation", () => {
  const data = new Uint8Array([1, 2, 3]);
  const name = "Logo";
  const description = "Company logo";
  const title = "Logo title";
  const hidden = true;

  const pptxSrc: PptxPictureOptions = {
    type: "png",
    data,
    width: 10,
    height: 10,
    name,
    description,
    title,
    hidden,
  };
  const xlsxSrc: XlsxPictureOptions = {
    type: "png",
    data,
    col: 1,
    row: 1,
    name,
    description,
    title,
    hidden,
  };
  const docxSrc: DocxPictureOptions = {
    type: "png",
    data,
    transformation: { width: 10, height: 10 },
    altText: { name, description, title, hidden },
  };

  it("pptx → docx: cNvPr lands on altText", () => {
    expect(toDocxPicture(pptxSrc).altText).toEqual({ name, description, title, hidden });
  });

  it("xlsx → docx: cNvPr lands on altText", () => {
    expect(toDocxPicture(xlsxSrc).altText).toEqual({ name, description, title, hidden });
  });

  it("docx → pptx: altText flows to top-level cNvPr", () => {
    const pptx = toPptxPicture(docxSrc);
    expect(pptx.name).toBe(name);
    expect(pptx.description).toBe(description);
    expect(pptx.title).toBe(title);
    expect(pptx.hidden).toBe(hidden);
  });

  it("docx → xlsx: altText flows to top-level cNvPr", () => {
    const xlsx = toXlsxPicture(docxSrc);
    expect(xlsx.name).toBe(name);
    expect(xlsx.description).toBe(description);
    expect(xlsx.title).toBe(title);
    expect(xlsx.hidden).toBe(hidden);
  });

  it("pptx → xlsx: cNvPr passes straight through", () => {
    const xlsx = toXlsxPicture(pptxSrc);
    expect(xlsx.name).toBe(name);
    expect(xlsx.description).toBe(description);
    expect(xlsx.title).toBe(title);
    expect(xlsx.hidden).toBe(hidden);
  });

  it("xlsx → pptx: cNvPr passes straight through", () => {
    const pptx = toPptxPicture(xlsxSrc);
    expect(pptx.name).toBe(name);
    expect(pptx.description).toBe(description);
    expect(pptx.title).toBe(title);
    expect(pptx.hidden).toBe(hidden);
  });

  it("docx defaults altText.name to Picture when source has other cNvPr but no name", () => {
    const docx = toDocxPicture({ ...pptxSrc, name: undefined });
    expect(docx.altText?.name).toBe("Picture");
    expect(docx.altText?.description).toBe(description);
  });

  it("omits docx altText when source carries no cNvPr at all", () => {
    const pptx: PptxPictureOptions = { type: "png", data, width: 10, height: 10 };
    expect(toDocxPicture(pptx).altText).toBeUndefined();
  });
});
