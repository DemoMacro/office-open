import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import {
  stringify,
  parse,
  type CustomDescriptor,
  type ReadContext,
  type WriteContext,
} from "../../descriptor";
import {
  diagramRelationshipIdsDesc,
  diagramStyleDesc,
  presentationLayoutVariablesDesc,
  diagramExtensionListDesc,
} from "./diagram-descriptors";
import type { DiagramExtensionListOptions } from "./diagram-props";
import type { DiagramRelationshipIdsOptions } from "./diagram-rel";
import type { DiagramStyleOptions } from "./diagram-style";
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
      extensions: [{ uri: "ext1" }, { uri: "ext2" }],
    };
    const result = roundTrip(diagramExtensionListDesc, opts);
    expect(result.extensions).toHaveLength(2);
    expect(result.extensions?.[0]?.uri).toBe("ext1");
    expect(result.extensions?.[1]?.uri).toBe("ext2");
  });

  it("returns undefined extensions when empty", () => {
    const opts: DiagramExtensionListOptions = {};
    const xml = stringify(diagramExtensionListDesc, opts, {} as WriteContext);
    expect(xml).toBeUndefined();
  });
});
