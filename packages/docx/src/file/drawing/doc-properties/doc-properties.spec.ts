import type { Context } from "@file/xml-components";
import { describe, expect, it } from "vite-plus/test";

import { DocProperties } from "./doc-properties";

const createMockContext = (): Context =>
  ({
    stack: [],
    viewWrapper: {
      relationships: {
        addRelationship: () => {},
      },
    },
  }) as never;

describe("DocProperties", () => {
  it("should add hlinkClick when hyperlink.click is specified", () => {
    const dp = new DocProperties({
      id: "1",
      name: "test",
      hyperlink: { click: "https://example.com" },
    });
    const xml = dp.toXml(createMockContext());
    expect(xml).to.contain("a:hlinkClick");
    expect(xml).to.contain('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
  });

  it("should add hlinkHover when hyperlink.hover is specified", () => {
    const dp = new DocProperties({
      id: "1",
      name: "test",
      hyperlink: { hover: "https://example.com/hover" },
    });
    const xml = dp.toXml(createMockContext());
    expect(xml).to.contain("a:hlinkHover");
  });

  it("should add both hlinkClick and hlinkHover when both are specified", () => {
    const dp = new DocProperties({
      id: "1",
      name: "test",
      hyperlink: {
        click: "https://example.com",
        hover: "https://example.com/hover",
      },
    });
    const xml = dp.toXml(createMockContext());
    expect(xml).to.contain("a:hlinkClick");
    expect(xml).to.contain("a:hlinkHover");
  });

  it("should register relationship for explicit click hyperlink", () => {
    let addedRelationship = false;
    const context: Context = {
      stack: [],
      viewWrapper: {
        relationships: {
          addRelationship: () => {
            addedRelationship = true;
          },
        },
      },
    } as never;

    const dp = new DocProperties({
      id: "1",
      name: "test",
      hyperlink: { click: "https://example.com" },
    });
    dp.toXml(context);
    expect(addedRelationship).toBe(true);
  });

  it("should register relationship for explicit hover hyperlink", () => {
    let addedCount = 0;
    const context: Context = {
      stack: [],
      viewWrapper: {
        relationships: {
          addRelationship: () => {
            addedCount++;
          },
        },
      },
    } as never;

    const dp = new DocProperties({
      id: "1",
      name: "test",
      hyperlink: {
        click: "https://example.com",
        hover: "https://example.com/hover",
      },
    });
    dp.toXml(context);
    expect(addedCount).toBe(2);
  });
});
