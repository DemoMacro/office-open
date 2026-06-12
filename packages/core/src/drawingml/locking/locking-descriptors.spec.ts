import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import type {
  ShapeLockingOptions,
  PictureLockingOptions,
  GroupLockingOptions,
  GraphicFrameLockingOptions,
} from "./locking";
import {
  shapeLockingDesc,
  pictureLockingDesc,
  groupLockingDesc,
  graphicFrameLockingDesc,
} from "./locking-descriptors";

function roundTrip<T>(desc: any, opts: T): T {
  const xml = stringify(desc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(desc, el, {} as any);
}

describe("shapeLockingDesc", () => {
  it("round-trips shape locking with all properties", () => {
    const opts: ShapeLockingOptions = {
      noGrp: true,
      noSelect: true,
      noRot: false,
      noChangeAspect: true,
      noMove: false,
      noResize: true,
      noTextEdit: true,
    };
    const result = roundTrip(shapeLockingDesc, opts);
    expect(result.noGrp).toBe(true);
    expect(result.noSelect).toBe(true);
    expect(result.noRot).toBe(false);
    expect(result.noChangeAspect).toBe(true);
    expect(result.noMove).toBe(false);
    expect(result.noResize).toBe(true);
    expect(result.noTextEdit).toBe(true);
  });

  it("round-trips minimal shape locking", () => {
    const opts: ShapeLockingOptions = { noGrp: true };
    const result = roundTrip(shapeLockingDesc, opts);
    expect(result.noGrp).toBe(true);
  });
});

describe("pictureLockingDesc", () => {
  it("round-trips picture locking with crop", () => {
    const opts: PictureLockingOptions = {
      noGrp: true,
      noCrop: true,
      noChangeAspect: true,
    };
    const result = roundTrip(pictureLockingDesc, opts);
    expect(result.noGrp).toBe(true);
    expect(result.noCrop).toBe(true);
    expect(result.noChangeAspect).toBe(true);
  });
});

describe("groupLockingDesc", () => {
  it("round-trips group locking", () => {
    const opts: GroupLockingOptions = {
      noGrp: true,
      noUngrp: true,
      noSelect: false,
    };
    const result = roundTrip(groupLockingDesc, opts);
    expect(result.noGrp).toBe(true);
    expect(result.noUngrp).toBe(true);
    expect(result.noSelect).toBe(false);
  });
});

describe("graphicFrameLockingDesc", () => {
  it("round-trips graphic frame locking", () => {
    const opts: GraphicFrameLockingOptions = {
      noGrp: true,
      noDrilldown: true,
      noSelect: false,
    };
    const result = roundTrip(graphicFrameLockingDesc, opts);
    expect(result.noGrp).toBe(true);
    expect(result.noDrilldown).toBe(true);
    expect(result.noSelect).toBe(false);
  });
});
