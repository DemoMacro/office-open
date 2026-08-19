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
      parts.push(`${key}="${opts[key] ? 1 : 0}"`);
    }
  }
  return parts.length ? " " + parts.join(" ") : "";
}

function readLockingAttrs(el: XmlElement, keys: readonly string[]): Record<string, boolean> {
  const result: Record<string, boolean> = {};
  if (!el.attributes) return result;
  for (const key of keys) {
    const val = el.attributes[key];
    if (val !== undefined) result[key] = parseOnOff(val) ?? false;
  }
  return result;
}

/** One locking descriptor: element tag plus the base keys and any extras. */
function createLockingDesc<T extends BaseLockingOptions>(
  tag: string,
  extraKeys: readonly string[] = [],
): CustomDescriptor<T> {
  const allKeys = [...BASE_LOCKING_KEYS, ...extraKeys];
  return {
    kind: "custom",
    stringify(opts, _ctx) {
      const attrStr = stringifyLockingAttrs(
        opts as unknown as Record<string, boolean | undefined>,
        allKeys,
      );
      // An empty <a:*Locks/> container is meaningful (locks present, none set)
      // and round-trips as {} — only an absent container stays undefined.
      if (!attrStr) return `<${tag}/>`;
      return `<${tag}${attrStr}/>`;
    },
    parse(el, _ctx) {
      return readLockingAttrs(el, allKeys) as unknown as T;
    },
  };
}

// ── Locking descriptors ──

export const shapeLockingDesc: CustomDescriptor<ShapeLockingOptions> =
  createLockingDesc<ShapeLockingOptions>("a:spLocks", SHAPE_EXTRA_KEYS);

export const pictureLockingDesc: CustomDescriptor<PictureLockingOptions> =
  createLockingDesc<PictureLockingOptions>("a:picLocks", PICTURE_EXTRA_KEYS);

export const groupLockingDesc: CustomDescriptor<GroupLockingOptions> =
  createLockingDesc<GroupLockingOptions>("a:grpSpLocks", GROUP_EXTRA_KEYS);

export const graphicFrameLockingDesc: CustomDescriptor<GraphicFrameLockingOptions> =
  createLockingDesc<GraphicFrameLockingOptions>("a:graphicFrameLocks", FRAME_EXTRA_KEYS);

// CT_ConnectorLocking = AG_Locking (base keys only)
export const connectorLockingDesc: CustomDescriptor<ConnectorLockingOptions> =
  createLockingDesc<ConnectorLockingOptions>("a:cxnSpLocks");
