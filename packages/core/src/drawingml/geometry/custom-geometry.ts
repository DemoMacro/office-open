/**
 * Custom geometry module for DrawingML shapes.
 *
 * Provides CT_CustomGeometry2D — user-defined 2D geometry with paths,
 * guides, adjust handles, connection sites, and a text insertion rectangle.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_CustomGeometry2D
 *
 * @module
 */
import { element } from "@office-open/xml";

import { xsdPathFillMode } from "../../util/mappings";
import type { GeometryGuide } from "./adjustment-values";

/**
 * Creates a geometry guide list element (a:avLst or a:gdLst).
 *
 * Both CT_GeomGuideList types share the same structure — a list of `a:gd` children.
 * The only difference is the wrapper element name.
 */
const createGuideList = (name: string, guides: readonly GeometryGuide[]): string => {
  const children = guides.map((guide) => `<a:gd name="${guide.name}" fmla="${guide.formula}"/>`);
  return element(name, undefined, children);
};

// ─── Path Fill Mode ─────────────────────────────────────────────────────────

/**
 * Path fill mode (ST_PathFillMode).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_PathFillMode">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="norm"/>
 *     <xsd:enumeration value="lighten"/>
 *     <xsd:enumeration value="lightenLess"/>
 *     <xsd:enumeration value="darken"/>
 *     <xsd:enumeration value="darkenLess"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const PathFillMode = {
  NONE: "none",
  NORM: "normal",
  LIGHTEN: "lighten",
  LIGHTEN_LESS: "lightenLess",
  DARKEN: "darken",
  DARKEN_LESS: "darkenLess",
} as const;

// ─── Point ──────────────────────────────────────────────────────────────────

/**
 * An adjustable point (CT_AdjPoint2D).
 *
 * Coordinates can be absolute values or references to geometry guide names.
 */
export interface AdjPoint {
  /** X coordinate (absolute value or guide name) */
  x: string;
  /** Y coordinate (absolute value or guide name) */
  y: string;
}

const createAdjPoint = (name: string, point: AdjPoint): string =>
  element(name, undefined, [`<a:pt x="${point.x}" y="${point.y}"/>`]);

// ─── Path Commands ──────────────────────────────────────────────────────────

/** Move-to path command (CT_Path2DMoveTo). */
export interface PathMoveTo {
  command: "moveTo";
  point: AdjPoint;
}

/** Line-to path command (CT_Path2DLineTo). */
export interface PathLineTo {
  command: "lineTo";
  point: AdjPoint;
}

/** Arc-to path command (CT_Path2DArcTo). */
export interface PathArcTo {
  command: "arcTo";
  widthRadius: string;
  heightRadius: string;
  startAngle: string;
  sweepAngle: string;
}

/** Quadratic Bezier-to path command (CT_Path2DQuadBezierTo). */
export interface PathQuadBezTo {
  command: "quadBezTo";
  points: readonly [AdjPoint, AdjPoint];
}

/** Cubic Bezier-to path command (CT_Path2DCubicBezierTo). */
export interface PathCubicBezTo {
  command: "cubicBezTo";
  points: readonly [AdjPoint, AdjPoint, AdjPoint];
}

/** Close path command (CT_Path2DClose). */
export interface PathClose {
  command: "close";
}

/** A path command within a custom geometry path. */
export type PathCommand =
  | PathMoveTo
  | PathLineTo
  | PathArcTo
  | PathQuadBezTo
  | PathCubicBezTo
  | PathClose;

const createPathCommand = (cmd: PathCommand): string => {
  switch (cmd.command) {
    case "moveTo":
      return createAdjPoint("a:moveTo", cmd.point);
    case "lineTo":
      return createAdjPoint("a:lnTo", cmd.point);
    case "arcTo":
      return `<a:arcTo wR="${cmd.widthRadius}" hR="${cmd.heightRadius}" stAng="${cmd.startAngle}" swAng="${cmd.sweepAngle}"/>`;
    case "quadBezTo":
      return element(
        "a:quadBezTo",
        undefined,
        cmd.points.map((pt) => `<a:pt x="${pt.x}" y="${pt.y}"/>`),
      );
    case "cubicBezTo":
      return element(
        "a:cubicBezTo",
        undefined,
        cmd.points.map((pt) => `<a:pt x="${pt.x}" y="${pt.y}"/>`),
      );
    case "close":
      return "<a:close/>";
  }
};

// ─── Path ───────────────────────────────────────────────────────────────────

/**
 * Options for a 2D path (CT_Path2D).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Path2D">
 *   <xsd:choice minOccurs="0" maxOccurs="unbounded">
 *     <xsd:element name="close" type="CT_Path2DClose"/>
 *     <xsd:element name="moveTo" type="CT_Path2DMoveTo"/>
 *     <xsd:element name="lnTo" type="CT_Path2DLineTo"/>
 *     <xsd:element name="arcTo" type="CT_Path2DArcTo"/>
 *     <xsd:element name="quadBezTo" type="CT_Path2DQuadBezierTo"/>
 *     <xsd:element name="cubicBezTo" type="CT_Path2DCubicBezierTo"/>
 *   </xsd:choice>
 *   <xsd:attribute name="w" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="h" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="fill" type="ST_PathFillMode" default="norm"/>
 *   <xsd:attribute name="stroke" type="xsd:boolean" default="true"/>
 *   <xsd:attribute name="extrusionOk" type="xsd:boolean" default="true"/>
 * </xsd:complexType>
 * ```
 */
export interface PathOptions {
  /** Path width (coordinate value, default 0 = use shape width) */
  w?: number;
  /** Path height (coordinate value, default 0 = use shape height) */
  h?: number;
  /** Fill mode for this path */
  fill?: (typeof PathFillMode)[keyof typeof PathFillMode];
  /** Whether to stroke the path */
  stroke?: boolean;
  /** Whether extrusion is allowed */
  extrusionOk?: boolean;
  /** Path commands (moveTo, lineTo, arcTo, etc.) */
  commands: readonly PathCommand[];
}

const createPath = (options: PathOptions): string => {
  const attrs: Record<string, string | number | boolean | undefined> = {};
  if (options.w !== undefined) attrs.w = options.w;
  if (options.h !== undefined) attrs.h = options.h;
  if (options.fill !== undefined) attrs.fill = xsdPathFillMode.to(options.fill);
  if (options.stroke !== undefined) attrs.stroke = options.stroke;
  if (options.extrusionOk !== undefined) attrs.extrusionOk = options.extrusionOk;

  return element("a:path", attrs, options.commands.map(createPathCommand));
};

// ─── Adjust Handle ──────────────────────────────────────────────────────────

/**
 * A position reference for adjust handles and connection sites.
 */
export interface AdjustHandlePosition {
  x: string;
  y: string;
}

/**
 * XY adjust handle (CT_XYAdjustHandle).
 *
 * Allows the user to drag a handle in X/Y directions.
 */
export interface XYAdjustHandle {
  type: "xy";
  guideRefX?: string;
  minX?: string;
  maxX?: string;
  guideRefY?: string;
  minY?: string;
  maxY?: string;
  position: AdjustHandlePosition;
}

/**
 * Polar adjust handle (CT_PolarAdjustHandle).
 *
 * Allows the user to drag a handle in radius/angle directions.
 */
export interface PolarAdjustHandle {
  type: "polar";
  guideRefRadius?: string;
  minRadius?: string;
  maxRadius?: string;
  guideRefAngle?: string;
  minAngle?: string;
  maxAngle?: string;
  position: AdjustHandlePosition;
}

/** An adjust handle (XY or Polar). */
export type AdjustHandle = XYAdjustHandle | PolarAdjustHandle;

const createAdjustHandlePos = (position: AdjustHandlePosition): string =>
  `<a:pos x="${position.x}" y="${position.y}"/>`;

const createXYAdjustHandle = (handle: XYAdjustHandle): string => {
  const attrs: Record<string, string | undefined> = {};
  if (handle.guideRefX !== undefined) attrs.gdRefX = handle.guideRefX;
  if (handle.minX !== undefined) attrs.minX = handle.minX;
  if (handle.maxX !== undefined) attrs.maxX = handle.maxX;
  if (handle.guideRefY !== undefined) attrs.gdRefY = handle.guideRefY;
  if (handle.minY !== undefined) attrs.minY = handle.minY;
  if (handle.maxY !== undefined) attrs.maxY = handle.maxY;

  return element("a:ahXY", attrs, [createAdjustHandlePos(handle.position)]);
};

const createPolarAdjustHandle = (handle: PolarAdjustHandle): string => {
  const attrs: Record<string, string | undefined> = {};
  if (handle.guideRefRadius !== undefined) attrs.gdRefR = handle.guideRefRadius;
  if (handle.minRadius !== undefined) attrs.minR = handle.minRadius;
  if (handle.maxRadius !== undefined) attrs.maxR = handle.maxRadius;
  if (handle.guideRefAngle !== undefined) attrs.gdRefAng = handle.guideRefAngle;
  if (handle.minAngle !== undefined) attrs.minAng = handle.minAngle;
  if (handle.maxAngle !== undefined) attrs.maxAng = handle.maxAngle;

  return element("a:ahPolar", attrs, [createAdjustHandlePos(handle.position)]);
};

// ─── Connection Site ────────────────────────────────────────────────────────

/**
 * Connection site (CT_ConnectionSite).
 *
 * Defines a point on the shape boundary where connectors can attach.
 */
export interface ConnectionSite {
  /** Angle (absolute value or guide name) */
  angle: string;
  /** Position */
  position: AdjustHandlePosition;
}

const createConnectionSite = (site: ConnectionSite): string =>
  element("a:cxn", { ang: site.angle }, [createAdjustHandlePos(site.position)]);

// ─── Geometry Rectangle ─────────────────────────────────────────────────────

/**
 * Text insertion rectangle (CT_GeomRect).
 *
 * Defines the rectangle within the shape where text is placed.
 * Coordinates can be absolute values or references to geometry guide names.
 */
export interface GeomRect {
  left: string;
  top: string;
  right: string;
  bottom: string;
}

const createGeomRect = (rect: GeomRect): string =>
  `<a:rect l="${rect.left}" t="${rect.top}" r="${rect.right}" b="${rect.bottom}"/>`;

// ─── Custom Geometry ────────────────────────────────────────────────────────

/**
 * Options for a custom 2D geometry (CT_CustomGeometry2D).
 *
 * Only `pathList` is required. All other sub-elements are optional.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_CustomGeometry2D">
 *   <xsd:sequence>
 *     <xsd:element name="avLst" type="CT_GeomGuideList" minOccurs="0"/>
 *     <xsd:element name="gdLst" type="CT_GeomGuideList" minOccurs="0"/>
 *     <xsd:element name="ahLst" type="CT_AdjustHandleList" minOccurs="0"/>
 *     <xsd:element name="cxnLst" type="CT_ConnectionSiteList" minOccurs="0"/>
 *     <xsd:element name="rect" type="CT_GeomRect" minOccurs="0"/>
 *     <xsd:element name="pathLst" type="CT_Path2DList" minOccurs="1"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Simple triangle
 * createCustomGeometry({
 *   pathList: [{
 *     commands: [
 *       { command: "moveTo", point: { x: "0", y: "0" } },
 *       { command: "lineTo", point: { x: "10000000", y: "0" } },
 *       { command: "lineTo", point: { x: "5000000", y: "10000000" } },
 *       { command: "close" },
 *     ],
 *   }],
 * });
 * ```
 */
export interface CustomGeometryOptions {
  /** Adjustment value guides (a:avLst) */
  adjustmentValues?: readonly GeometryGuide[];
  /** Geometry guide formulas (a:gdLst) */
  guides?: readonly GeometryGuide[];
  /** Adjust handles (a:ahLst) */
  adjustHandles?: readonly AdjustHandle[];
  /** Connection sites (a:cxnLst) */
  connectionSites?: readonly ConnectionSite[];
  /** Text insertion rectangle (a:rect) */
  textRectangle?: GeomRect;
  /** Path definitions (a:pathLst) — required */
  pathList: readonly PathOptions[];
}

/**
 * Creates a custom 2D geometry element as an XML string (a:custGeom).
 *
 * @example
 * ```typescript
 * // Diamond shape
 * const geom = createCustomGeometry({
 *   pathList: [{
 *     commands: [
 *       { command: "moveTo", point: { x: "5000000", y: "0" } },
 *       { command: "lineTo", point: { x: "10000000", y: "5000000" } },
 *       { command: "lineTo", point: { x: "5000000", y: "10000000" } },
 *       { command: "lineTo", point: { x: "0", y: "5000000" } },
 *       { command: "close" },
 *     ],
 *   }],
 *   textRectangle: { left: "2000000", top: "2000000", right: "8000000", bottom: "8000000" },
 * });
 * ```
 */
export const createCustomGeometry = (options: CustomGeometryOptions): string => {
  const children: string[] = [];

  // a:avLst
  if (options.adjustmentValues) {
    children.push(createGuideList("a:avLst", options.adjustmentValues));
  }

  // a:gdLst
  if (options.guides) {
    children.push(createGuideList("a:gdLst", options.guides));
  }

  // a:ahLst
  if (options.adjustHandles && options.adjustHandles.length > 0) {
    children.push(
      element(
        "a:ahLst",
        undefined,
        options.adjustHandles.map((h) =>
          h.type === "xy" ? createXYAdjustHandle(h) : createPolarAdjustHandle(h),
        ),
      ),
    );
  }

  // a:cxnLst
  if (options.connectionSites && options.connectionSites.length > 0) {
    children.push(
      element("a:cxnLst", undefined, options.connectionSites.map(createConnectionSite)),
    );
  }

  // a:rect
  if (options.textRectangle) {
    children.push(createGeomRect(options.textRectangle));
  }

  // a:pathLst (required)
  children.push(element("a:pathLst", undefined, options.pathList.map(createPath)));

  return element("a:custGeom", undefined, children);
};
