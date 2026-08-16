/**
 * Non-visual content part properties (a:CT_NonVisualContentPartProperties).
 *
 * Carries the cpLocks locking set and the isComment flag for wp:contentPart
 * wrappers (CT_WordprocessingContentPart). The cNvContentPartPr element itself
 * is host-namespaced (wp:/wpg:/wpc:), so the tag is caller-supplied.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_NonVisualContentPartProperties
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import { parseOnOff } from "../../util/values";
import { createContentPartLocking } from "../locking/locking";
import type { ContentPartLockingOptions } from "../locking/locking";

export interface NonVisualContentPartPropertiesOptions {
  /** Locking flags (a:cpLocks — CT_ContentPartLocking) */
  locks?: ContentPartLockingOptions;
  /**
   * Whether the content part hosts a comment (`@isComment`, default true).
   * Only `false` deviates from the default and is emitted explicitly.
   */
  isComment?: boolean;
}

/** Locking attribute keys shared by stringify/parse. */
const LOCKING_KEYS: readonly (keyof ContentPartLockingOptions & string)[] = [
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

/**
 * Stringify a cNvContentPartPr element (CT_NonVisualContentPartProperties).
 * The `isComment` default is true, so only `false` is written.
 */
export function stringifyNonVisualContentPartProperties(
  tag: string,
  opts: NonVisualContentPartPropertiesOptions | undefined,
): string {
  if (!opts) return `<${tag}/>`;
  let inner = "";
  if (opts.locks) inner += createContentPartLocking(opts.locks);
  const attr = opts.isComment === false ? ' isComment="0"' : "";
  return inner ? `<${tag}${attr}>${inner}</${tag}>` : `<${tag}${attr}/>`;
}

/** Parse a cNvContentPartPr element into its options form. */
export function parseNonVisualContentPartProperties(
  el: XmlElement | undefined,
): NonVisualContentPartPropertiesOptions | undefined {
  if (!el) return undefined;
  const result: NonVisualContentPartPropertiesOptions = {};
  const locks = findChild(el, "a:cpLocks");
  if (locks) {
    const flags: Record<string, boolean> = {};
    for (const key of LOCKING_KEYS) {
      const raw = locks.attributes?.[key];
      if (raw !== undefined) flags[key] = parseOnOff(raw) ?? false;
    }
    if (Object.keys(flags).length > 0) result.locks = flags as ContentPartLockingOptions;
  }
  const isComment = el.attributes?.["isComment"];
  if (isComment !== undefined) result.isComment = parseOnOff(isComment) ?? true;
  return Object.keys(result).length > 0 ? result : undefined;
}
