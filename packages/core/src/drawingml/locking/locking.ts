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
  noGrp?: boolean;
  /** Prevent selection */
  noSelect?: boolean;
  /** Prevent rotation */
  noRot?: boolean;
  /** Prevent aspect ratio change */
  noChangeAspect?: boolean;
  /** Prevent moving */
  noMove?: boolean;
  /** Prevent resizing */
  noResize?: boolean;
  /** Prevent edit points */
  noEditPoints?: boolean;
  /** Prevent adjust handles */
  noAdjustHandles?: boolean;
  /** Prevent changing arrowheads */
  noChangeArrowheads?: boolean;
  /** Prevent changing shape type */
  noChangeShapeType?: boolean;
}

// ── Shape locking (CT_ShapeLocking) ──

export interface ShapeLockingOptions extends BaseLockingOptions {
  /** Prevent text editing */
  noTextEdit?: boolean;
}

// ── Picture locking (CT_PictureLocking) ──

export interface PictureLockingOptions extends BaseLockingOptions {
  /** Prevent cropping */
  noCrop?: boolean;
}

// ── Group locking (CT_GroupLocking) ──

export interface GroupLockingOptions {
  noGrp?: boolean;
  noUngrp?: boolean;
  noSelect?: boolean;
  noRot?: boolean;
  noChangeAspect?: boolean;
  noMove?: boolean;
  noResize?: boolean;
}

// ── Graphic frame locking (CT_GraphicalObjectFrameLocking) ──

export interface GraphicFrameLockingOptions {
  noGrp?: boolean;
  noDrilldown?: boolean;
  noSelect?: boolean;
  noChangeAspect?: boolean;
  noMove?: boolean;
  noResize?: boolean;
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
