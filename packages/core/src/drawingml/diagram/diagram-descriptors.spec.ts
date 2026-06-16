import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import {
  diagramRelIdsDesc,
  diagramStyleDesc,
  presLayoutVarsDesc,
  diagramExtLstDesc,
} from "./diagram-descriptors";
import type { DiagramExtLstOptions } from "./diagram-props";
import type { DiagramRelIdsOptions } from "./diagram-rel";
import type { DiagramStyleOptions } from "./diagram-style";
import type { PresLayoutVarsOptions } from "./layout-vars";

function roundTrip<T>(desc: any, opts: T): T {
  const xml = stringify(desc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(desc, el, {} as any);
}

describe("diagramRelIdsDesc", () => {
  it("round-trips all relationship IDs", () => {
    const opts: DiagramRelIdsOptions = {
      dm: "rId1",
      lo: "rId2",
      qs: "rId3",
      cs: "rId4",
    };
    const result = roundTrip(diagramRelIdsDesc, opts);
    expect(result.dm).toBe("rId1");
    expect(result.lo).toBe("rId2");
    expect(result.qs).toBe("rId3");
    expect(result.cs).toBe("rId4");
  });
});

describe("diagramStyleDesc", () => {
  it("round-trips style indices", () => {
    const opts: DiagramStyleOptions = {
      lineReference: { idx: 2 },
      fillReference: { idx: 3 },
      effectReference: { idx: 1 },
      fontReference: { idx: "major" },
    };
    const result = roundTrip(diagramStyleDesc, opts);
    expect(result.lineReference?.idx).toBe(2);
    expect(result.fillReference?.idx).toBe(3);
    expect(result.effectReference?.idx).toBe(1);
    expect(result.fontReference?.idx).toBe("major");
  });

  it("round-trips with defaults", () => {
    const opts: DiagramStyleOptions = {};
    const result = roundTrip(diagramStyleDesc, opts);
    // Default values from stringify
    expect(result.lineReference?.idx).toBe(1);
    expect(result.fillReference?.idx).toBe(1);
    expect(result.effectReference?.idx).toBe(0);
    expect(result.fontReference?.idx).toBe("minor");
  });
});

describe("presLayoutVarsDesc", () => {
  it("round-trips org chart flag", () => {
    const opts: PresLayoutVarsOptions = {
      orgChart: { val: true },
    };
    const result = roundTrip(presLayoutVarsDesc, opts);
    expect(result.orgChart?.val).toBe(true);
  });

  it("round-trips hierarchy branch", () => {
    const opts: PresLayoutVarsOptions = {
      hierBranch: { val: "hang" },
      chMax: { val: 4 },
      chPref: { val: 2 },
    };
    const result = roundTrip(presLayoutVarsDesc, opts);
    expect(result.hierBranch?.val).toBe("hang");
    expect(result.chMax?.val).toBe(4);
    expect(result.chPref?.val).toBe(2);
  });

  it("round-trips animation options", () => {
    const opts: PresLayoutVarsOptions = {
      animOne: { val: "one" },
      animLvl: { val: "lvl" },
    };
    const result = roundTrip(presLayoutVarsDesc, opts);
    expect(result.animOne?.val).toBe("one");
    expect(result.animLvl?.val).toBe("lvl");
  });
});

describe("diagramExtLstDesc", () => {
  it("round-trips extensions", () => {
    const opts: DiagramExtLstOptions = {
      extensions: [{ uri: "ext1" }, { uri: "ext2" }],
    };
    const result = roundTrip(diagramExtLstDesc, opts);
    expect(result.extensions).toHaveLength(2);
    expect(result.extensions?.[0].uri).toBe("ext1");
    expect(result.extensions?.[1].uri).toBe("ext2");
  });

  it("returns undefined extensions when empty", () => {
    const opts: DiagramExtLstOptions = {};
    const xml = stringify(diagramExtLstDesc, opts, {} as any);
    expect(xml).toBeUndefined();
  });
});
