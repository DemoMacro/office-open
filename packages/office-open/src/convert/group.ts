/**
 * Cross-format group conversion.
 *
 * Groups convert between pptx (p:grpSp), xlsx (xdr:grpSp), and docx (wpg). The
 * group container (bounding box + rotation/flip) round-trips near-losslessly
 * via the absolute-box model (./position); only the xlsx leg loses precise
 * positioning (heuristic cell anchors).
 *
 * Children recurse through their own converters — shapes via ./shape,
 * connectors via ./connector. docx has no standalone connector and xlsx groups
 * hold only shapes/connectors, so picture/table/chart/... children are dropped
 * with a warning on the legs that cannot host them.
 *
 * @module
 */

import type { ShapePropertiesOptions, GroupTransform2DOptions } from "@office-open/core/drawingml";
import type {
  WpgGroupRunOptions as DocxGroupOptions,
  GroupChildMediaData,
  MediaDataTransformation,
  WpsShapeCoreOptions,
} from "@office-open/docx";
import { createTransformation } from "@office-open/docx";
import type {
  GroupShapeOptions as PptxGroupOptions,
  SlideChild,
  ShapeOptions as PptxShapeOptions,
  ConnectorShapeOptions as PptxConnectorOptions,
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

const EMU_PER_PIXEL = 9525;
const ANGLE_UNITS_PER_DEGREE = 60_000;

// ── container helpers ──

/** docx child MediaDataTransformation → absolute box (reads EMUs; falls back to pixels). */
function docxChildMediaToBox(t: MediaDataTransformation): AbsoluteBox {
  const x = t.offset?.emus?.x ?? (t.offset?.pixels.x ?? 0) * EMU_PER_PIXEL;
  const y = t.offset?.emus?.y ?? (t.offset?.pixels.y ?? 0) * EMU_PER_PIXEL;
  return {
    x,
    y,
    width: t.emus.x,
    height: t.emus.y,
    ...(t.rotation !== undefined ? { rotation: t.rotation / ANGLE_UNITS_PER_DEGREE } : {}),
    ...(t.flip?.horizontal ? { flipHorizontal: true } : {}),
    ...(t.flip?.vertical ? { flipVertical: true } : {}),
  };
}

// ── shape-child spPr bridge ──

/** pptx shape top-level → core spPr (group-child position is absolute). */
function pptxShapeToSpPr(shape: PptxShapeOptions): ShapePropertiesOptions {
  return {
    x: shape.x,
    y: shape.y,
    width: shape.width,
    height: shape.height,
    ...(shape.rotation !== undefined ? { rotation: shape.rotation } : {}),
    ...(shape.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(shape.geometry !== undefined ? { geometry: shape.geometry } : {}),
    ...(shape.customGeometry !== undefined ? { customGeometry: shape.customGeometry } : {}),
    ...pickContent(shape),
  };
}

/** core spPr → pptx shape top-level fields. */
function spPrToPptxShape(spPr: ShapePropertiesOptions): PptxShapeOptions {
  return {
    x: spPr.x,
    y: spPr.y,
    width: spPr.width,
    height: spPr.height,
    ...(spPr.rotation !== undefined ? { rotation: spPr.rotation } : {}),
    ...(spPr.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(spPr.flipVertical ? { flipVertical: true } : {}),
    ...(spPr.geometry !== undefined ? { geometry: spPr.geometry } : {}),
    ...(spPr.customGeometry !== undefined ? { customGeometry: spPr.customGeometry } : {}),
    ...pickContent(spPr),
  };
}

/** xlsx group child shape → docx wps core (position lives on the wpg child wrapper). */
function xlsxShapeChildToDocxData(s: GroupShapeChildOptions): WpsShapeCoreOptions {
  const preset = toPresetGeometry(s.spPr.geometry);
  return {
    children: s.textBody ? textBodyToDocxChildren(s.textBody) : [],
    ...pickContent(s.spPr),
    ...(s.spPr.customGeometry !== undefined ? { customGeometry: s.spPr.customGeometry } : {}),
    ...(preset !== undefined ? { presetGeometry: preset } : {}),
    ...(s.name ? { nonVisualProperties: { name: s.name } } : {}),
  };
}

/** docx wps child → core spPr (absolute position from the child transformation). */
function docxChildToSpPr(data: WpsShapeCoreOptions, box: AbsoluteBox): ShapePropertiesOptions {
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
  if (data.presetGeometry !== undefined) out.geometry = data.presetGeometry;
  else if (data.customGeometry !== undefined) out.customGeometry = data.customGeometry;
  return out;
}

/** xlsx group connector child → pptx connector. */
function xlsxConnectorChildToPptx(c: GroupConnectorChildOptions): PptxConnectorOptions {
  const { x1, y1, x2, y2 } = boxToEndpoints(boxFromSpPr(c.spPr));
  return {
    x1,
    y1,
    x2,
    y2,
    ...(c.spPr.outline !== undefined ? { outline: c.spPr.outline } : {}),
    ...(c.spPr.fill !== undefined ? { fill: c.spPr.fill } : {}),
    ...(c.locking ? { locking: c.locking } : {}),
    ...(c.startConnection ? { startConnection: c.startConnection } : {}),
    ...(c.endConnection ? { endConnection: c.endConnection } : {}),
    ...(c.name ? { name: c.name } : {}),
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
  if ("grpSpPr" in source) {
    const g = source.grpSpPr;
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
  return { children, transformation: boxToDocx(box) };
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
      transformation: createTransformation(boxToDocx(boxFromSpPr(s.spPr))),
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
  if ("grpSpPr" in source) {
    const g = source.grpSpPr;
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
  return { ...boxToPptx(box), children };
}

function xlsxGroupChildrenToPptx(
  shapes: GroupShapeChildOptions[] | undefined,
  connectors: GroupConnectorChildOptions[] | undefined,
): SlideChild[] {
  const out: SlideChild[] = [];
  for (const s of shapes ?? []) {
    const shape = spPrToPptxShape(s.spPr);
    if (s.textBody) shape.textBody = s.textBody;
    if (s.name) shape.name = s.name;
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
      if (child.data.nonVisualProperties?.name) shape.name = child.data.nonVisualProperties.name;
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
    grpSpPr,
    ...(shapes.length ? { shapes } : {}),
    ...(connectors.length ? { connectors } : {}),
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
        spPr: pptxShapeToSpPr(child.shape),
        ...(child.shape.textBody ? { textBody: child.shape.textBody } : {}),
        ...(child.shape.name ? { name: child.shape.name } : {}),
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
    ...(c.outline !== undefined ? { outline: c.outline } : {}),
    ...(c.fill !== undefined ? { fill: c.fill } : {}),
    ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(box.flipVertical ? { flipVertical: true } : {}),
  };
  return {
    spPr,
    ...(c.locking ? { locking: c.locking } : {}),
    ...(c.startConnection ? { startConnection: c.startConnection } : {}),
    ...(c.endConnection ? { endConnection: c.endConnection } : {}),
    ...(c.name ? { name: c.name } : {}),
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
        spPr,
        ...(textBody ? { textBody } : {}),
        ...(child.data.nonVisualProperties?.name
          ? { name: child.data.nonVisualProperties.name }
          : {}),
      });
    } else {
      console.warn(`Unsupported docx group child → xlsx (${child.type}); skipped.`);
    }
  }
  return { shapes, connectors };
}
