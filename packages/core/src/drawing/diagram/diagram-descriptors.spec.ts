import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import {
  stringify,
  parse,
  type CustomDescriptor,
  type ReadContext,
  type WriteContext,
} from "../../descriptor";
import type { DefaultShapeStyleOptions } from "../../theme/theme-options";
import {
  diagramRelationshipIdsDesc,
  diagramStyleDesc,
  presentationLayoutVariablesDesc,
  diagramExtensionListDesc,
} from "./diagram-descriptors";
import type { DiagramExtensionListOptions } from "./diagram-props";
import type { DiagramRelationshipIdsOptions } from "./diagram-rel";
import {
  colorsDefinitionHeaderListDesc,
  layoutDefinitionHeaderListDesc,
  styleDefinitionHeaderListDesc,
} from "./headers";
import type {
  ColorsDefinitionHeaderListOptions,
  LayoutDefinitionHeaderListOptions,
  StyleDefinitionHeaderListOptions,
} from "./headers";
import type { PresentationLayoutVariablesOptions } from "./layout-vars";

function roundTrip<T>(desc: CustomDescriptor<T>, opts: T): T {
  const xml = stringify(desc, opts, {} as WriteContext);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parse(desc, el, {} as ReadContext);
}

describe("diagramRelationshipIdsDesc", () => {
  it("round-trips all relationship IDs", () => {
    const opts: DiagramRelationshipIdsOptions = {
      dataModel: "rId1",
      layout: "rId2",
      quickStyle: "rId3",
      colorStyle: "rId4",
    };
    const result = roundTrip(diagramRelationshipIdsDesc, opts);
    expect(result.dataModel).toBe("rId1");
    expect(result.layout).toBe("rId2");
    expect(result.quickStyle).toBe("rId3");
    expect(result.colorStyle).toBe("rId4");
  });
});

describe("diagramStyleDesc", () => {
  it("round-trips style indices", () => {
    const opts: DefaultShapeStyleOptions = {
      lineReference: { index: 2, color: { value: "accent2" } },
      fillReference: { index: 3, color: { value: "FF0000" } },
      effectReference: { index: 1 },
      fontReference: { collection: "major", color: { value: "tx2" } },
    };
    const result = roundTrip(diagramStyleDesc, opts);
    expect(result.lineReference?.index).toBe(2);
    expect(result.lineReference?.color).toEqual({ value: "accent2" });
    expect(result.fillReference?.index).toBe(3);
    expect(result.fillReference?.color).toEqual({ value: "FF0000" });
    expect(result.effectReference?.index).toBe(1);
    expect(result.fontReference?.collection).toBe("major");
    expect(result.fontReference?.color).toEqual({ value: "tx2" });
  });
});

describe("presentationLayoutVariablesDesc", () => {
  it("round-trips org chart flag", () => {
    const opts: PresentationLayoutVariablesOptions = {
      orgChart: { val: true },
    };
    const result = roundTrip(presentationLayoutVariablesDesc, opts);
    expect(result.orgChart?.val).toBe(true);
  });

  it("round-trips hierarchy branch", () => {
    const opts: PresentationLayoutVariablesOptions = {
      hierBranch: { val: "hanging" },
      maxChildren: { val: 4 },
      preferredChildren: { val: 2 },
    };
    const result = roundTrip(presentationLayoutVariablesDesc, opts);
    expect(result.hierBranch?.val).toBe("hanging");
    expect(result.maxChildren?.val).toBe(4);
    expect(result.preferredChildren?.val).toBe(2);
  });

  it("round-trips bullet, direction, and resize-handle variables", () => {
    const opts: PresentationLayoutVariablesOptions = {
      bulletEnabled: { val: true },
      direction: { val: "reversed" },
      resizeHandles: { val: "relative" },
    };
    const result = roundTrip(presentationLayoutVariablesDesc, opts);
    expect(result.bulletEnabled?.val).toBe(true);
    expect(result.direction?.val).toBe("reversed");
    expect(result.resizeHandles?.val).toBe("relative");
  });

  it("round-trips animation options", () => {
    const opts: PresentationLayoutVariablesOptions = {
      animateOneByOne: { val: "one" },
      animationLevel: { val: "level" },
    };
    const result = roundTrip(presentationLayoutVariablesDesc, opts);
    expect(result.animateOneByOne?.val).toBe("one");
    expect(result.animationLevel?.val).toBe("level");
  });
});

describe("diagramExtensionListDesc", () => {
  it("round-trips extensions", () => {
    const opts: DiagramExtensionListOptions = {
      extensions: [{ uri: "ext1" }, { uri: "ext2", content: '<foo:payload bar="1"/>' }],
    };
    const result = roundTrip(diagramExtensionListDesc, opts);
    expect(result.extensions).toHaveLength(2);
    expect(result.extensions?.[0]?.uri).toBe("ext1");
    expect(result.extensions?.[1]?.uri).toBe("ext2");
    expect(result.extensions?.[1]?.content).toBe('<foo:payload bar="1"/>');
  });

  it("returns undefined extensions when empty", () => {
    const opts: DiagramExtensionListOptions = {};
    const xml = stringify(diagramExtensionListDesc, opts, {} as WriteContext);
    expect(xml).toBeUndefined();
  });
});

describe("diagram definition header lists", () => {
  const category = { type: "layout", pri: 25000 };
  const extensions = [{ uri: "header-extension", content: '<foo:payload bar="1"/>' }];

  it("round-trips all colors definition headers", () => {
    const opts: ColorsDefinitionHeaderListOptions = {
      headers: [
        {
          uniqueId: "color-header",
          minVer: "12.0",
          resId: 2,
          title: [{ lang: "en-US", val: "Color" }],
          desc: [{ lang: "en-US", val: "Color description" }],
          categories: [category],
          extensions,
        },
      ],
    };
    const result = roundTrip(colorsDefinitionHeaderListDesc, opts);
    expect(result.headers?.[0]?.uniqueId).toBe("color-header");
    expect(result.headers?.[0]?.minVer).toBe("12.0");
    expect(result.headers?.[0]?.resId).toBe(2);
    expect(result.headers?.[0]?.title).toEqual([{ lang: "en-US", val: "Color" }]);
    expect(result.headers?.[0]?.desc).toEqual([{ lang: "en-US", val: "Color description" }]);
    expect(result.headers?.[0]?.categories).toEqual([category]);
    expect(result.headers?.[0]?.extensions).toEqual(extensions);
  });

  it("round-trips all layout definition headers", () => {
    const opts: LayoutDefinitionHeaderListOptions = {
      headers: [
        {
          uniqueId: "layout-header",
          defStyle: "layout-style",
          resId: 3,
          title: [{ val: "Layout" }],
          desc: [{ val: "Layout description" }],
          categories: [category],
          extensions,
        },
      ],
    };
    const result = roundTrip(layoutDefinitionHeaderListDesc, opts);
    expect(result.headers?.[0]).toEqual(opts.headers?.[0]);
  });

  it("round-trips all style definition headers", () => {
    const opts: StyleDefinitionHeaderListOptions = {
      headers: [
        {
          uniqueId: "style-header",
          minVer: "12.0",
          resId: 4,
          title: [{ val: "Style" }],
          desc: [{ val: "Style description" }],
          categories: [category],
          extensions,
        },
      ],
    };
    const result = roundTrip(styleDefinitionHeaderListDesc, opts);
    expect(result.headers?.[0]).toEqual(opts.headers?.[0]);
  });
});
