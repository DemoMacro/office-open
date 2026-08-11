import { describe, expect, it } from "vite-plus/test";

import type { LayoutDefinition, MasterDefinition } from "./file";
import { PLACEHOLDER_TYPE_TO_KEY, resolvePlaceholder } from "./placeholder";

describe("PLACEHOLDER_TYPE_TO_KEY", () => {
  it("maps standard p:ph/@type tokens (ctrTitle normalizes to title)", () => {
    expect(PLACEHOLDER_TYPE_TO_KEY.title).toBe("title");
    expect(PLACEHOLDER_TYPE_TO_KEY.ctrTitle).toBe("title");
    expect(PLACEHOLDER_TYPE_TO_KEY.body).toBe("body");
    expect(PLACEHOLDER_TYPE_TO_KEY.sub).toBe("subtitle");
    expect(PLACEHOLDER_TYPE_TO_KEY.dt).toBe("date");
    expect(PLACEHOLDER_TYPE_TO_KEY.ftr).toBe("footer");
    expect(PLACEHOLDER_TYPE_TO_KEY.sldNum).toBe("slideNumber");
  });
});

describe("resolvePlaceholder", () => {
  it("inherits position from layout", () => {
    const layout: LayoutDefinition = {
      placeholders: { title: { x: 1, y: 2, width: 3, height: 4 } },
    };
    expect(resolvePlaceholder("title", layout, undefined).position).toEqual({
      x: 1,
      y: 2,
      width: 3,
      height: 4,
    });
  });

  it("inherits position from master when layout has none", () => {
    const master: MasterDefinition = {
      placeholders: { title: { x: 10, y: 20, width: 30, height: 40 } },
    };
    expect(resolvePlaceholder("title", undefined, master).position?.x).toBe(10);
  });

  it("layout position takes precedence over master", () => {
    const layout: LayoutDefinition = {
      placeholders: { title: { x: 1, y: 2, width: 3, height: 4 } },
    };
    const master: MasterDefinition = {
      placeholders: { title: { x: 10, y: 20, width: 30, height: 40 } },
    };
    expect(resolvePlaceholder("title", layout, master).position?.x).toBe(1);
  });

  it('reports hidden when layout carries sz="0" (false)', () => {
    const layout: LayoutDefinition = { placeholders: { date: false } };
    expect(resolvePlaceholder("dt", layout, undefined).hidden).toBe(true);
  });

  it("reports hidden when master carries false", () => {
    const master: MasterDefinition = { placeholders: { footer: false } };
    expect(resolvePlaceholder("ftr", undefined, master).hidden).toBe(true);
  });

  it("ctrTitle normalizes to the title slot", () => {
    const layout: LayoutDefinition = {
      placeholders: { title: { x: 5, y: 6, width: 7, height: 8 } },
    };
    expect(resolvePlaceholder("ctrTitle", layout, undefined).position?.x).toBe(5);
  });

  it("returns empty for unknown placeholder type", () => {
    expect(resolvePlaceholder("unknown", undefined, undefined)).toEqual({});
  });

  it("returns empty when master placeholder is boolean (no explicit position)", () => {
    // Master fresh API allows `true` (use reference position) which carries no
    // concrete coordinates — resolvePlaceholder cannot synthesize them.
    const master: MasterDefinition = { placeholders: { title: true } };
    expect(resolvePlaceholder("title", undefined, master)).toEqual({});
  });
});

describe("resolvePlaceholder facets", () => {
  it("inherits geometry facet from layout", () => {
    const layout: LayoutDefinition = {
      placeholders: { title: { x: 1, y: 2, width: 3, height: 4, geometry: "ellipse" } },
    };
    expect(resolvePlaceholder("title", layout, undefined).facets?.geometry).toBe("ellipse");
  });

  it("inherits style facet from master", () => {
    const master: MasterDefinition = {
      placeholders: {
        title: { x: 1, y: 2, width: 3, height: 4, style: { fillReference: { index: 2 } } },
      },
    };
    expect(resolvePlaceholder("title", undefined, master).facets?.style?.fillReference?.index).toBe(
      2,
    );
  });

  it("layout facets override master per-facet; master fills gaps layout omits", () => {
    const layout: LayoutDefinition = {
      placeholders: { title: { x: 1, y: 2, width: 3, height: 4, geometry: "ellipse" } },
    };
    const master: MasterDefinition = {
      placeholders: {
        title: {
          x: 10,
          y: 20,
          width: 30,
          height: 40,
          geometry: "rect",
          style: { fillReference: { index: 1 } },
        },
      },
    };
    const result = resolvePlaceholder("title", layout, master);
    expect(result.position?.x).toBe(1);
    expect(result.facets?.geometry).toBe("ellipse");
    expect(result.facets?.style?.fillReference?.index).toBe(1);
  });

  it("returns no facets when neither layer defines them (position only)", () => {
    const layout: LayoutDefinition = {
      placeholders: { title: { x: 1, y: 2, width: 3, height: 4 } },
    };
    expect(resolvePlaceholder("title", layout, undefined).facets).toBeUndefined();
  });
});
