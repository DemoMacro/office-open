/**
 * Cross-format group conversion.
 *
 * Groups convert between pptx (p:grpSp), xlsx (xdr:grpSp), and docx (wpg). The
 * group container (bounding box + rotation/flip) round-trips near-losslessly
 * via the absolute-box model (./position); only the xlsx leg loses precise
 * positioning (heuristic cell anchors).
 *
 * The container cNvPr (name/description/title/hidden) passes straight through on
 * every leg — pptx/xlsx via pickGroupBase, docx via its altText bridge — so alt
 * text survives a cross-format copy. Child shapes/connectors carry their own
 * cNvPr through pickNonVisualDrawingProperties (all four fields, not just name).
 *
 * Children recurse through their own converters — shapes via ./shape,
 * connectors via ./connector. docx has no standalone connector and xlsx groups
 * hold only shapes/connectors, so picture/table/chart/... children are dropped
 * with a warning on the legs that cannot host them.
 *
 * @module
 */

import {
  convertPixelsToEmu,
  parseAngle,
  pickGroupBase,
  pickNonVisualDrawingProperties,
} from "@office-open/core";
import type { NonVisualDrawingPropertiesOptions } from "@office-open/core";
import type { ShapePropertiesOptions, GroupTransform2DOptions } from "@office-open/core/drawing";
import type {
  GroupOptions as DocxGroupOptions,
  GroupChildMediaData,
  MediaDataTransformation,
  ShapeCoreOptions,
} from "@office-open/docx";
import { createTransformation } from "@office-open/docx";
import type {
  GroupOptions as PptxGroupOptions,
  SlideChild,
  ShapeOptions as PptxShapeOptions,
  ConnectorOptions as PptxConnectorOptions,
} from "@office-open/pptx";
import type {
  GroupOptions as XlsxGroupOptions,
  GroupShapeChildOptions,
  GroupConnectorChildOptions,
} from "@office-open/xlsx";

import { boxToEndpoints, endpointsToBox } from "./connector";
import {
  boxFromPptx,
  boxFromSpPr,
  boxFromXlsxAnchor,
  boxFromDocx,
  boxToPptx,
  boxToXlsx,
  boxToDocx,
} from "./position";
import type { AbsoluteBox } from "./position";
import {
  docxToTextBody,
  pickContent,
  textBodyToDocxChildren,
  toDocxShapeParts,
  toPresetGeometry,
} from "./shape";

// ── container cNvPr bridge ──

/**
 * Build the docx altText (wp:docPr) from the container cNvPr. Only emitted when
 * at least one cNvPr field is authored; name defaults to "Group" since docx
 * requires it. Structurally compatible with docx's DocPropertiesOptions without
 * importing that internal type.
 */
const altTextFromCnvPr = (
  picked: Partial<NonVisualDrawingPropertiesOptions>,
): { altText?: NonVisualDrawingPropertiesOptions & { name: string } } => {
  if (
    picked.name === undefined &&
    picked.description === undefined &&
    picked.title === undefined &&
    picked.hidden === undefined
  ) {
    return {};
  }
  return { altText: { name: picked.name ?? "Group", ...picked } };
};

/**
 * Build a docx child nonVisualProperties object from a picked cNvPr. Only
 * emitted when at least one field is authored; name defaults to `fallbackName`.
 */
const docxNonVisualFromCnvPr = (
  picked: Partial<NonVisualDrawingPropertiesOptions>,
  fallbackName: string,
): { nonVisualProperties: NonVisualDrawingPropertiesOptions } => {
  const name = picked.name ?? fallbackName;
  return { nonVisualProperties: { name, ...picked } };
};

/** True when a picked cNvPr carries at least one authored field. */
const hasCnvPr = (picked: Partial<NonVisualDrawingPropertiesOptions>): boolean =>
  picked.name !== undefined ||
  picked.description !== undefined ||
  picked.title !== undefined ||
  picked.hidden !== undefined;

// ── container helpers ──

/** docx child MediaDataTransformation → absolute box (reads EMUs; falls back to pixels). */
function docxChildMediaToBox(t: MediaDataTransformation): AbsoluteBox {
  const x = t.offset?.emus?.x ?? convertPixelsToEmu(t.offset?.pixels.x ?? 0);
  const y = t.offset?.emus?.y ?? convertPixelsToEmu(t.offset?.pixels.y ?? 0);
  return {
    x,
    y,
    width: t.emus.x,
    height: t.emus.y,
    ...(t.rotation !== undefined ? { rotation: parseAngle(t.rotation) } : {}),
    ...(t.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(t.flipVertical ? { flipVertical: true } : {}),
  };
}

// ── shape-child spPr bridge ──

/** pptx shape → core spPr (group-child position is absolute). */
function pptxShapeToSpPr(shape: PptxShapeOptions): ShapePropertiesOptions {
  const paint = shape.properties ?? {};
  return {
    x: shape.x,
    y: shape.y,
    width: shape.width,
    height: shape.height,
    ...(shape.rotation !== undefined ? { rotation: shape.rotation } : {}),
    ...(shape.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(paint.geometry !== undefined ? { geometry: paint.geometry } : {}),
    ...(paint.customGeometry !== undefined ? { customGeometry: paint.customGeometry } : {}),
    ...pickContent(paint),
  };
}

/** core spPr → pptx shape (transform top-level, paint nested in `properties`). */
function spPrToPptxShape(spPr: ShapePropertiesOptions): PptxShapeOptions {
  return {
    x: spPr.x,
    y: spPr.y,
    width: spPr.width,
    height: spPr.height,
    ...(spPr.rotation !== undefined ? { rotation: spPr.rotation } : {}),
    ...(spPr.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(spPr.flipVertical ? { flipVertical: true } : {}),
    properties: {
      ...(spPr.geometry !== undefined ? { geometry: spPr.geometry } : {}),
      ...(spPr.customGeometry !== undefined ? { customGeometry: spPr.customGeometry } : {}),
      ...pickContent(spPr),
    },
  };
}

/** xlsx group child shape → docx wps core (position lives on the wpg child wrapper). */
function xlsxShapeChildToDocxData(s: GroupShapeChildOptions): ShapeCoreOptions {
  const preset = toPresetGeometry(s.properties.geometry);
  const cnvPr = pickNonVisualDrawingProperties(s);
  return {
    children: s.textBody ? textBodyToDocxChildren(s.textBody) : [],
    ...pickContent(s.properties),
    ...(s.properties.customGeometry !== undefined
      ? { customGeometry: s.properties.customGeometry }
      : {}),
    ...(preset !== undefined ? { geometry: preset } : {}),
    ...(hasCnvPr(cnvPr) ? docxNonVisualFromCnvPr(cnvPr, "Shape") : {}),
  };
}

/** docx wps child → core spPr (absolute position from the child transformation). */
function docxChildToSpPr(data: ShapeCoreOptions, box: AbsoluteBox): ShapePropertiesOptions {
  const out: ShapePropertiesOptions = {
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    ...(box.rotation !== undefined ? { rotation: box.rotation } : {}),
    ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(box.flipVertical ? { flipVertical: true } : {}),
    ...pickContent(data),
  };
  if (data.geometry !== undefined) out.geometry = data.geometry;
  else if (data.customGeometry !== undefined) out.customGeometry = data.customGeometry;
  return out;
}

/** xlsx group connector child → pptx connector. */
function xlsxConnectorChildToPptx(c: GroupConnectorChildOptions): PptxConnectorOptions {
  const { x1, y1, x2, y2 } = boxToEndpoints(boxFromSpPr(c.properties));
  return {
    x1,
    y1,
    x2,
    y2,
    properties: pickContent(c.properties),
    ...(c.locking ? { locking: c.locking } : {}),
    ...(c.startConnection ? { startConnection: c.startConnection } : {}),
    ...(c.endConnection ? { endConnection: c.endConnection } : {}),
    ...pickNonVisualDrawingProperties(c),
  };
}

// ── → docx ──

/** Convert a pptx group to a docx wpg group. */
export function toDocxGroup(source: PptxGroupOptions): DocxGroupOptions;
/** Convert an xlsx group to a docx wpg group. */
export function toDocxGroup(source: XlsxGroupOptions): DocxGroupOptions;
export function toDocxGroup(source: PptxGroupOptions | XlsxGroupOptions): DocxGroupOptions {
  let box: AbsoluteBox;
  let children: GroupChildMediaData[];
  if ("properties" in source) {
    const g = source.properties;
    box = boxFromXlsxAnchor(
      source,
      g.width,
      g.height,
      g.rotation,
      g.flipHorizontal,
      g.flipVertical,
    );
    children = xlsxGroupChildrenToDocx(source.shapes, source.connectors);
  } else {
    box = boxFromPptx(
      source.x,
      source.y,
      source.width,
      source.height,
      source.rotation,
      source.flipHorizontal,
    );
    children = pptxGroupChildrenToDocx(source.children);
  }
  // Container cNvPr → docx altText (wp:docPr).
  return { children, transformation: boxToDocx(box), ...altTextFromCnvPr(pickGroupBase(source)) };
}

function pptxGroupChildrenToDocx(children: SlideChild[] | undefined): GroupChildMediaData[] {
  const out: GroupChildMediaData[] = [];
  for (const child of children ?? []) {
    if ("shape" in child) {
      const parts = toDocxShapeParts(child.shape);
      out.push({
        type: "wps",
        transformation: createTransformation(parts.transformation),
        data: parts.data,
      });
    } else if ("connector" in child) {
      console.warn("Connector in group → docx is unsupported; skipped.");
    } else {
      console.warn(`Unsupported group child → docx (${Object.keys(child)[0]}); skipped.`);
    }
  }
  return out;
}

function xlsxGroupChildrenToDocx(
  shapes: GroupShapeChildOptions[] | undefined,
  connectors: GroupConnectorChildOptions[] | undefined,
): GroupChildMediaData[] {
  const out: GroupChildMediaData[] = [];
  for (const s of shapes ?? []) {
    out.push({
      type: "wps",
      transformation: createTransformation(boxToDocx(boxFromSpPr(s.properties))),
      data: xlsxShapeChildToDocxData(s),
    });
  }
  if (connectors?.length) {
    console.warn("Connector in group → docx is unsupported; skipped.");
  }
  return out;
}

// ── → pptx ──

/** Convert a docx wpg group to a pptx group. */
export function toPptxGroup(source: DocxGroupOptions): PptxGroupOptions;
/** Convert an xlsx group to a pptx group. */
export function toPptxGroup(source: XlsxGroupOptions): PptxGroupOptions;
export function toPptxGroup(source: DocxGroupOptions | XlsxGroupOptions): PptxGroupOptions {
  let box: AbsoluteBox;
  let children: SlideChild[];
  // Container cNvPr: docx bridges through altText; xlsx extends BaseGroupOptions.
  const cnvPr =
    "transformation" in source
      ? pickNonVisualDrawingProperties(source.altText)
      : pickGroupBase(source);
  if ("properties" in source) {
    const g = source.properties;
    box = boxFromXlsxAnchor(
      source,
      g.width,
      g.height,
      g.rotation,
      g.flipHorizontal,
      g.flipVertical,
    );
    children = xlsxGroupChildrenToPptx(source.shapes, source.connectors);
  } else {
    box = boxFromDocx(source.transformation);
    children = docxGroupChildrenToPptx(source.children);
  }
  return { ...boxToPptx(box), children, ...cnvPr };
}

function xlsxGroupChildrenToPptx(
  shapes: GroupShapeChildOptions[] | undefined,
  connectors: GroupConnectorChildOptions[] | undefined,
): SlideChild[] {
  const out: SlideChild[] = [];
  for (const s of shapes ?? []) {
    const shape = spPrToPptxShape(s.properties);
    if (s.textBody) shape.textBody = s.textBody;
    Object.assign(shape, pickNonVisualDrawingProperties(s));
    out.push({ shape });
  }
  for (const c of connectors ?? []) {
    out.push({ connector: xlsxConnectorChildToPptx(c) });
  }
  return out;
}

function docxGroupChildrenToPptx(children: GroupChildMediaData[] | undefined): SlideChild[] {
  const out: SlideChild[] = [];
  for (const child of children ?? []) {
    if (child.type === "wps") {
      const box = docxChildMediaToBox(child.transformation);
      const shape = spPrToPptxShape(docxChildToSpPr(child.data, box));
      const textBody = docxToTextBody(child.data.children, child.data.bodyProperties);
      if (textBody) shape.textBody = textBody;
      Object.assign(shape, pickNonVisualDrawingProperties(child.data.nonVisualProperties));
      out.push({ shape });
    } else {
      console.warn(`Unsupported docx group child → pptx (${child.type}); skipped.`);
    }
  }
  return out;
}

// ── → xlsx ──

/** Convert a docx wpg group to an xlsx group. */
export function toXlsxGroup(source: DocxGroupOptions): XlsxGroupOptions;
/** Convert a pptx group to an xlsx group. */
export function toXlsxGroup(source: PptxGroupOptions): XlsxGroupOptions;
export function toXlsxGroup(source: DocxGroupOptions | PptxGroupOptions): XlsxGroupOptions {
  let box: AbsoluteBox;
  let shapes: GroupShapeChildOptions[];
  let connectors: GroupConnectorChildOptions[];
  // Container cNvPr: docx bridges through altText; pptx extends BaseGroupOptions.
  const cnvPr =
    "transformation" in source
      ? pickNonVisualDrawingProperties(source.altText)
      : pickGroupBase(source);
  if ("transformation" in source) {
    box = boxFromDocx(source.transformation);
    const r = docxGroupChildrenToXlsx(source.children);
    shapes = r.shapes;
    connectors = r.connectors;
  } else {
    box = boxFromPptx(
      source.x,
      source.y,
      source.width,
      source.height,
      source.rotation,
      source.flipHorizontal,
    );
    const r = pptxGroupChildrenToXlsx(source.children);
    shapes = r.shapes;
    connectors = r.connectors;
  }
  const pos = boxToXlsx(box);
  const grpSpPr: GroupTransform2DOptions = {
    x: pos.xfrmX,
    y: pos.xfrmY,
    width: box.width,
    height: box.height,
    ...(box.rotation !== undefined ? { rotation: box.rotation } : {}),
    ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(box.flipVertical ? { flipVertical: true } : {}),
  };
  return {
    ...pos.anchor,
    properties: grpSpPr,
    ...(shapes.length ? { shapes } : {}),
    ...(connectors.length ? { connectors } : {}),
    ...cnvPr,
  };
}

function pptxGroupChildrenToXlsx(children: SlideChild[] | undefined): {
  shapes: GroupShapeChildOptions[];
  connectors: GroupConnectorChildOptions[];
} {
  const shapes: GroupShapeChildOptions[] = [];
  const connectors: GroupConnectorChildOptions[] = [];
  for (const child of children ?? []) {
    if ("shape" in child) {
      const s: GroupShapeChildOptions = {
        properties: pptxShapeToSpPr(child.shape),
        ...(child.shape.textBody ? { textBody: child.shape.textBody } : {}),
        ...pickNonVisualDrawingProperties(child.shape),
      };
      shapes.push(s);
    } else if ("connector" in child) {
      connectors.push(pptxConnectorToXlsxChild(child.connector));
    } else {
      console.warn(`Unsupported group child → xlsx (${Object.keys(child)[0]}); skipped.`);
    }
  }
  return { shapes, connectors };
}

function pptxConnectorToXlsxChild(c: PptxConnectorOptions): GroupConnectorChildOptions {
  // Place the connector's endpoint box on spPr.xfrm (no anchor — group children
  // position via spPr). The shared helper encodes direction as flip flags.
  const box = endpointsToBox(c.x1, c.y1, c.x2, c.y2);
  const spPr: ShapePropertiesOptions = {
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    geometry: "line",
    ...pickContent(c.properties ?? {}),
    ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(box.flipVertical ? { flipVertical: true } : {}),
  };
  return {
    properties: spPr,
    ...(c.locking ? { locking: c.locking } : {}),
    ...(c.startConnection ? { startConnection: c.startConnection } : {}),
    ...(c.endConnection ? { endConnection: c.endConnection } : {}),
    ...pickNonVisualDrawingProperties(c),
  };
}

function docxGroupChildrenToXlsx(children: GroupChildMediaData[] | undefined): {
  shapes: GroupShapeChildOptions[];
  connectors: GroupConnectorChildOptions[];
} {
  const shapes: GroupShapeChildOptions[] = [];
  const connectors: GroupConnectorChildOptions[] = [];
  for (const child of children ?? []) {
    if (child.type === "wps") {
      const box = docxChildMediaToBox(child.transformation);
      const spPr = docxChildToSpPr(child.data, box);
      const textBody = docxToTextBody(child.data.children, child.data.bodyProperties);
      shapes.push({
        properties: spPr,
        ...(textBody ? { textBody } : {}),
        ...pickNonVisualDrawingProperties(child.data.nonVisualProperties),
      });
    } else {
      console.warn(`Unsupported docx group child → xlsx (${child.type}); skipped.`);
    }
  }
  return { shapes, connectors };
}
