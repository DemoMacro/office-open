/**
 * Locking descriptor for DrawingML shapes.
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { parseOnOff } from "../../util/values";
import type {
  BaseLockingOptions,
  ConnectorLockingOptions,
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

const SHAPE_LOCKING_KEYS: readonly (keyof ShapeLockingOptions & string)[] = [
  ...BASE_LOCKING_KEYS,
  "noTextEdit",
];
const PICTURE_LOCKING_KEYS: readonly (keyof PictureLockingOptions & string)[] = [
  ...BASE_LOCKING_KEYS,
  "noCrop",
];
const GROUP_LOCKING_KEYS: readonly (keyof GroupLockingOptions & string)[] = [
  "noGrp",
  "noUngrp",
  "noSelect",
  "noRot",
  "noChangeAspect",
  "noMove",
  "noResize",
];
const FRAME_LOCKING_KEYS: readonly (keyof GraphicFrameLockingOptions & string)[] = [
  "noGrp",
  "noDrilldown",
  "noSelect",
  "noChangeAspect",
  "noMove",
  "noResize",
];

type LockingOptions = BaseLockingOptions | GroupLockingOptions | GraphicFrameLockingOptions;

// ── Helper: stringify boolean attributes ──

function stringifyLockingAttrs<T extends LockingOptions>(
  opts: T,
  keys: readonly (keyof T & string)[],
): string {
  const parts: string[] = [];
  for (const key of keys) {
    if (opts[key] !== undefined) {
      parts.push(`${key}="${opts[key] ? 1 : 0}"`);
    }
  }
  return parts.length ? " " + parts.join(" ") : "";
}

function readLockingAttrs<T extends LockingOptions, K extends keyof T & string>(
  el: XmlElement,
  keys: readonly K[],
): T {
  const result: Partial<Record<K, boolean>> = {};
  const attrs = el.attributes;
  if (!attrs) return result as T;
  for (const key of keys) {
    const val = attrs[key];
    if (val !== undefined) result[key] = parseOnOff(val) ?? false;
  }
  return result as T;
}

/** One locking descriptor: element tag plus the base keys and any extras. */
function createLockingDesc<T extends LockingOptions>(
  tag: string,
  keys: readonly (keyof T & string)[],
): CustomDescriptor<T> {
  return {
    kind: "custom",
    stringify(opts, _ctx) {
      const attrStr = stringifyLockingAttrs(opts, keys);
      // An empty <a:*Locks/> container is meaningful (locks present, none set)
      // and round-trips as {} — only an absent container stays undefined.
      if (!attrStr) return `<${tag}/>`;
      return `<${tag}${attrStr}/>`;
    },
    parse(el, _ctx) {
      return readLockingAttrs(el, keys);
    },
  };
}

// ── Locking descriptors ──

export const shapeLockingDesc: CustomDescriptor<ShapeLockingOptions> =
  createLockingDesc<ShapeLockingOptions>("a:spLocks", SHAPE_LOCKING_KEYS);

export const pictureLockingDesc: CustomDescriptor<PictureLockingOptions> =
  createLockingDesc<PictureLockingOptions>("a:picLocks", PICTURE_LOCKING_KEYS);

export const groupLockingDesc: CustomDescriptor<GroupLockingOptions> =
  createLockingDesc<GroupLockingOptions>("a:grpSpLocks", GROUP_LOCKING_KEYS);

export const graphicFrameLockingDesc: CustomDescriptor<GraphicFrameLockingOptions> =
  createLockingDesc<GraphicFrameLockingOptions>("a:graphicFrameLocks", FRAME_LOCKING_KEYS);

// CT_ConnectorLocking = AG_Locking (base keys only)
export const connectorLockingDesc: CustomDescriptor<ConnectorLockingOptions> =
  createLockingDesc<ConnectorLockingOptions>("a:cxnSpLocks", BASE_LOCKING_KEYS);
