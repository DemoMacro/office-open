/**
 * Locking descriptor for DrawingML shapes.
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import type {
  BaseLockingOptions,
  GraphicFrameLockingOptions,
  GroupLockingOptions,
  PictureLockingOptions,
  ShapeLockingOptions,
} from "./locking";

// ── Common locking attribute keys ──

const BASE_LOCKING_KEYS: readonly (keyof BaseLockingOptions & string)[] = [
  "noGrp",
  "noSelect",
  "noRot",
  "noChangeAspect",
  "noMove",
  "noResize",
  "noEditPoints",
  "noAdjustHandles",
  "noChangeArrowheads",
  "noChangeShapeType",
];

const SHAPE_EXTRA_KEYS: readonly (keyof ShapeLockingOptions & string)[] = ["noTextEdit"];
const PICTURE_EXTRA_KEYS: readonly (keyof PictureLockingOptions & string)[] = ["noCrop"];
const GROUP_EXTRA_KEYS: readonly (keyof GroupLockingOptions & string)[] = ["noUngrp"];
const FRAME_EXTRA_KEYS: readonly (keyof GraphicFrameLockingOptions & string)[] = ["noDrilldown"];

// ── Helper: stringify boolean attributes ──

function stringifyLockingAttrs(
  opts: Readonly<Record<string, boolean | undefined>>,
  keys: readonly string[],
): string {
  const parts: string[] = [];
  for (const key of keys) {
    if (opts[key] !== undefined) {
      parts.push(`a:${key}="${opts[key] ? 1 : 0}"`);
    }
  }
  return parts.length ? " " + parts.join(" ") : "";
}

function readLockingAttrs(el: XmlElement, keys: readonly string[]): Record<string, boolean> {
  const result: Record<string, boolean> = {};
  if (!el.attributes) return result;
  for (const key of keys) {
    const val = el.attributes[`a:${key}`];
    if (val !== undefined) result[key] = val === "1" || val === 1 || val === "true";
  }
  return result;
}

// ── Shape locking descriptor ──

export const shapeLockingDesc: CustomDescriptor<ShapeLockingOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...SHAPE_EXTRA_KEYS];
    const attrStr = stringifyLockingAttrs(opts as Record<string, boolean | undefined>, allKeys);
    if (!attrStr) return undefined;
    return `<a:spLocks${attrStr}/>`;
  },
  parse(el, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...SHAPE_EXTRA_KEYS];
    return readLockingAttrs(el, allKeys) as Partial<ShapeLockingOptions>;
  },
};

// ── Picture locking descriptor ──

export const pictureLockingDesc: CustomDescriptor<PictureLockingOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...PICTURE_EXTRA_KEYS];
    const attrStr = stringifyLockingAttrs(opts as Record<string, boolean | undefined>, allKeys);
    if (!attrStr) return undefined;
    return `<a:picLocks${attrStr}/>`;
  },
  parse(el, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...PICTURE_EXTRA_KEYS];
    return readLockingAttrs(el, allKeys) as Partial<PictureLockingOptions>;
  },
};

// ── Group locking descriptor ──

export const groupLockingDesc: CustomDescriptor<GroupLockingOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...GROUP_EXTRA_KEYS];
    const attrStr = stringifyLockingAttrs(opts as Record<string, boolean | undefined>, allKeys);
    if (!attrStr) return undefined;
    return `<a:grpSpLocks${attrStr}/>`;
  },
  parse(el, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...GROUP_EXTRA_KEYS];
    return readLockingAttrs(el, allKeys) as Partial<GroupLockingOptions>;
  },
};

// ── Graphic frame locking descriptor ──

export const graphicFrameLockingDesc: CustomDescriptor<GraphicFrameLockingOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...FRAME_EXTRA_KEYS];
    const attrStr = stringifyLockingAttrs(opts as Record<string, boolean | undefined>, allKeys);
    if (!attrStr) return undefined;
    return `<a:graphicFrameLocks${attrStr}/>`;
  },
  parse(el, _ctx) {
    const allKeys = [...BASE_LOCKING_KEYS, ...FRAME_EXTRA_KEYS];
    return readLockingAttrs(el, allKeys) as Partial<GraphicFrameLockingOptions>;
  },
};
