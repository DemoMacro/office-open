import type { ConnectorOptions as PptxConnectorOptions } from "@office-open/pptx";
import type { ConnectorOptions as XlsxConnectorOptions } from "@office-open/xlsx";
import { describe, expect, it, vi } from "vitest";

import { toDocxConnector, toPptxConnector, toXlsxConnector } from "./connector";

describe("toXlsxConnector (pptx → xlsx)", () => {
  it("maps endpoints to box + cell anchor and emits prstGeom line", () => {
    const pptx: PptxConnectorOptions = {
      x1: 0,
      y1: 0,
      x2: 1219200,
      y2: 381000,
      outline: { type: "solidFill", color: { value: "000000" }, width: 9525 },
    };
    const x = toXlsxConnector(pptx);
    expect(x.col).toBe(1);
    expect(x.row).toBe(1);
    expect(x.toCol).toBe(3);
    expect(x.toRow).toBe(3);
    expect(x.properties.geometry).toBe("line");
    expect(x.properties.width).toBe(1219200);
    expect(x.properties.height).toBe(381000);
    expect(x.properties.outline?.color).toEqual({ value: "000000" });
  });

  it("encodes a reversed horizontal line as flipHorizontal", () => {
    const x = toXlsxConnector({ x1: 1219200, y1: 0, x2: 0, y2: 0 });
    expect(x.properties.flipHorizontal).toBe(true);
    expect(x.properties.width).toBe(1219200);
  });

  it("carries locking + endpoint glue verbatim", () => {
    const x = toXlsxConnector({
      x1: 0,
      y1: 0,
      x2: 100,
      y2: 100,
      locking: {},
      startConnection: { id: 1, index: 0 },
      endConnection: { id: 2, index: 1 },
    });
    expect(x.locking).toEqual({});
    expect(x.startConnection).toEqual({ id: 1, index: 0 });
    expect(x.endConnection).toEqual({ id: 2, index: 1 });
  });
});

describe("toPptxConnector (xlsx → pptx)", () => {
  it("restores endpoints from box, reversing flip", () => {
    const xlsx = toXlsxConnector({ x1: 1219200, y1: 0, x2: 0, y2: 0 });
    const back = toPptxConnector(xlsx);
    expect(back.x1).toBe(1219200);
    expect(back.x2).toBe(0);
  });
});

describe("toDocxConnector", () => {
  it("warns and returns undefined (docx has no standalone connector)", () => {
    const spy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const result = toDocxConnector({ x1: 0, y1: 0, x2: 100, y2: 100 });
    expect(result).toBeUndefined();
    expect(spy).toHaveBeenCalledOnce();
    spy.mockRestore();
  });
});

describe("round-trip pptx → xlsx → pptx", () => {
  it("preserves endpoints + outline", () => {
    const pptx: PptxConnectorOptions = {
      x1: 0,
      y1: 0,
      x2: 1219200,
      y2: 381000,
      outline: { type: "solidFill", color: { value: "000000" }, width: 9525 },
    };
    const back = toPptxConnector(toXlsxConnector(pptx));
    expect(back.x1).toBe(0);
    expect(back.y1).toBe(0);
    expect(back.x2).toBe(1219200);
    expect(back.y2).toBe(381000);
    expect(back.outline).toEqual({ type: "solidFill", color: { value: "000000" }, width: 9525 });
  });
});

describe("cross-format cNvPr + locking + endpoint preservation", () => {
  const name = "Flow arrow";
  const description = "Process flow arrow";
  const title = "Arrow title";
  const hidden = true;
  const locking = { noAdjustHandles: true };
  const startConnection = { id: 1, index: 0 };
  const endConnection = { id: 2, index: 1 };

  it("pptx → xlsx: base fields pass straight through", () => {
    const xlsx = toXlsxConnector({
      x1: 0,
      y1: 0,
      x2: 100,
      y2: 100,
      name,
      description,
      title,
      hidden,
      locking,
      startConnection,
      endConnection,
    });
    expect(xlsx.name).toBe(name);
    expect(xlsx.description).toBe(description);
    expect(xlsx.title).toBe(title);
    expect(xlsx.hidden).toBe(hidden);
    expect(xlsx.locking).toEqual(locking);
    expect(xlsx.startConnection).toEqual(startConnection);
    expect(xlsx.endConnection).toEqual(endConnection);
  });

  it("xlsx → pptx: base fields pass straight through", () => {
    const xlsx: XlsxConnectorOptions = {
      col: 1,
      row: 1,
      properties: { geometry: "line", width: 100, height: 100 },
      name,
      description,
      title,
      hidden,
      locking,
      startConnection,
      endConnection,
    };
    const pptx = toPptxConnector(xlsx);
    expect(pptx.name).toBe(name);
    expect(pptx.description).toBe(description);
    expect(pptx.title).toBe(title);
    expect(pptx.hidden).toBe(hidden);
    expect(pptx.locking).toEqual(locking);
    expect(pptx.startConnection).toEqual(startConnection);
    expect(pptx.endConnection).toEqual(endConnection);
  });
});
