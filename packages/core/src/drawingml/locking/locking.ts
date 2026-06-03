/**
 * DrawingML shape locking properties.
 *
 * Controls which aspects of a shape can be modified by users.
 *
 * Reference: OOXML transitional, dml-main.xsd, AG_Locking / CT_ShapeLocking / CT_PictureLocking / CT_GroupLocking
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";

// ── Common locking attributes (AG_Locking) ──

export interface BaseLockingOptions {
  /** Prevent grouping */
  readonly noGrp?: boolean;
  /** Prevent selection */
  readonly noSelect?: boolean;
  /** Prevent rotation */
  readonly noRot?: boolean;
  /** Prevent aspect ratio change */
  readonly noChangeAspect?: boolean;
  /** Prevent moving */
  readonly noMove?: boolean;
  /** Prevent resizing */
  readonly noResize?: boolean;
  /** Prevent edit points */
  readonly noEditPoints?: boolean;
  /** Prevent adjust handles */
  readonly noAdjustHandles?: boolean;
  /** Prevent changing arrowheads */
  readonly noChangeArrowheads?: boolean;
  /** Prevent changing shape type */
  readonly noChangeShapeType?: boolean;
}

// ── Shape locking (CT_ShapeLocking) ──

export interface ShapeLockingOptions extends BaseLockingOptions {
  /** Prevent text editing */
  readonly noTextEdit?: boolean;
}

// ── Picture locking (CT_PictureLocking) ──

export interface PictureLockingOptions extends BaseLockingOptions {
  /** Prevent cropping */
  readonly noCrop?: boolean;
}

// ── Group locking (CT_GroupLocking) ──

export interface GroupLockingOptions {
  readonly noGrp?: boolean;
  readonly noUngrp?: boolean;
  readonly noSelect?: boolean;
  readonly noRot?: boolean;
  readonly noChangeAspect?: boolean;
  readonly noMove?: boolean;
  readonly noResize?: boolean;
}

// ── Graphic frame locking (CT_GraphicalObjectFrameLocking) ──

export interface GraphicFrameLockingOptions {
  readonly noGrp?: boolean;
  readonly noDrilldown?: boolean;
  readonly noSelect?: boolean;
  readonly noChangeAspect?: boolean;
  readonly noMove?: boolean;
  readonly noResize?: boolean;
}

// ── Factory functions ──

function toAttrs(opts: Readonly<Record<string, boolean | undefined>>): {
  readonly [key: string]: { readonly key: string; readonly value: boolean | undefined };
} {
  const result: Record<string, { key: string; value: boolean | undefined }> = {};
  for (const key of Object.keys(opts)) {
    if (opts[key] !== undefined) {
      result[key] = { key: `a:${key}`, value: opts[key] };
    }
  }
  return result;
}

export function createShapeLocking(opts: ShapeLockingOptions): BuilderElement {
  return new BuilderElement({
    name: "a:spLocks",
    attributes: toAttrs(opts as Record<string, boolean | undefined>),
  });
}

export function createPictureLocking(opts: PictureLockingOptions): BuilderElement {
  return new BuilderElement({
    name: "a:picLocks",
    attributes: toAttrs(opts as Record<string, boolean | undefined>),
  });
}

export function createGroupLocking(opts: GroupLockingOptions): BuilderElement {
  return new BuilderElement({
    name: "a:grpSpLocks",
    attributes: toAttrs(opts as Record<string, boolean | undefined>),
  });
}

export function createGraphicFrameLocking(opts: GraphicFrameLockingOptions): BuilderElement {
  return new BuilderElement({
    name: "a:graphicFrameLocks",
    attributes: toAttrs(opts as Record<string, boolean | undefined>),
  });
}
