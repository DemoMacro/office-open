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
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { GeometryGuide } from "../preset-geometry/adjustment-values/adjustment-values";

/**
 * Creates a geometry guide list element (a:avLst or a:gdLst).
 *
 * Both CT_GeomGuideList types share the same structure — a list of `a:gd` children.
 * The only difference is the wrapper element name.
 */
const createGuideList = (name: string, guides: readonly GeometryGuide[]): XmlComponent => {
    const children: XmlComponent[] = guides.map(
        (guide) =>
            new BuilderElement<{ readonly name: string; readonly fmla: string }>({
                attributes: {
                    name: { key: "name", value: guide.name },
                    fmla: { key: "fmla", value: guide.formula },
                },
                name: "a:gd",
            }),
    );
    return new BuilderElement({ name, children });
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
    NORM: "norm",
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
    readonly x: string;
    /** Y coordinate (absolute value or guide name) */
    readonly y: string;
}

const createAdjPoint = (name: string, point: AdjPoint): XmlComponent =>
    new BuilderElement({
        name,
        children: [
            new BuilderElement<{ readonly x: string; readonly y: string }>({
                name: "a:pos",
                attributes: {
                    x: { key: "x", value: point.x },
                    y: { key: "y", value: point.y },
                },
            }),
        ],
    });

// ─── Path Commands ──────────────────────────────────────────────────────────

/** Move-to path command (CT_Path2DMoveTo). */
export interface PathMoveTo {
    readonly command: "moveTo";
    readonly point: AdjPoint;
}

/** Line-to path command (CT_Path2DLineTo). */
export interface PathLineTo {
    readonly command: "lineTo";
    readonly point: AdjPoint;
}

/** Arc-to path command (CT_Path2DArcTo). */
export interface PathArcTo {
    readonly command: "arcTo";
    readonly wR: string;
    readonly hR: string;
    readonly stAng: string;
    readonly swAng: string;
}

/** Quadratic Bezier-to path command (CT_Path2DQuadBezierTo). */
export interface PathQuadBezTo {
    readonly command: "quadBezTo";
    readonly points: readonly [AdjPoint, AdjPoint];
}

/** Cubic Bezier-to path command (CT_Path2DCubicBezierTo). */
export interface PathCubicBezTo {
    readonly command: "cubicBezTo";
    readonly points: readonly [AdjPoint, AdjPoint, AdjPoint];
}

/** Close path command (CT_Path2DClose). */
export interface PathClose {
    readonly command: "close";
}

/** A path command within a custom geometry path. */
export type PathCommand =
    | PathMoveTo
    | PathLineTo
    | PathArcTo
    | PathQuadBezTo
    | PathCubicBezTo
    | PathClose;

const createPathCommand = (cmd: PathCommand): XmlComponent => {
    switch (cmd.command) {
        case "moveTo":
            return createAdjPoint("a:moveTo", cmd.point);
        case "lineTo":
            return createAdjPoint("a:lnTo", cmd.point);
        case "arcTo":
            return new BuilderElement<{
                readonly wR: string;
                readonly hR: string;
                readonly stAng: string;
                readonly swAng: string;
            }>({
                name: "a:arcTo",
                attributes: {
                    wR: { key: "wR", value: cmd.wR },
                    hR: { key: "hR", value: cmd.hR },
                    stAng: { key: "stAng", value: cmd.stAng },
                    swAng: { key: "swAng", value: cmd.swAng },
                },
            });
        case "quadBezTo":
            return new BuilderElement({
                name: "a:quadBezTo",
                children: cmd.points.map(
                    (pt) =>
                        new BuilderElement<{ readonly x: string; readonly y: string }>({
                            name: "a:pos",
                            attributes: {
                                x: { key: "x", value: pt.x },
                                y: { key: "y", value: pt.y },
                            },
                        }),
                ),
            });
        case "cubicBezTo":
            return new BuilderElement({
                name: "a:cubicBezTo",
                children: cmd.points.map(
                    (pt) =>
                        new BuilderElement<{ readonly x: string; readonly y: string }>({
                            name: "a:pos",
                            attributes: {
                                x: { key: "x", value: pt.x },
                                y: { key: "y", value: pt.y },
                            },
                        }),
                ),
            });
        case "close":
            return new BuilderElement({ name: "a:close" });
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
    readonly w?: number;
    /** Path height (coordinate value, default 0 = use shape height) */
    readonly h?: number;
    /** Fill mode for this path */
    readonly fill?: (typeof PathFillMode)[keyof typeof PathFillMode];
    /** Whether to stroke the path */
    readonly stroke?: boolean;
    /** Whether extrusion is allowed */
    readonly extrusionOk?: boolean;
    /** Path commands (moveTo, lineTo, arcTo, etc.) */
    readonly commands: readonly PathCommand[];
}

const createPath = (options: PathOptions): XmlComponent => {
    const attrs: Record<
        string,
        { readonly key: string; readonly value: string | number | boolean }
    > = {};
    if (options.w !== undefined) attrs.w = { key: "w", value: options.w };
    if (options.h !== undefined) attrs.h = { key: "h", value: options.h };
    if (options.fill !== undefined) attrs.fill = { key: "fill", value: options.fill };
    if (options.stroke !== undefined) attrs.stroke = { key: "stroke", value: options.stroke };
    if (options.extrusionOk !== undefined)
        attrs.extrusionOk = { key: "extrusionOk", value: options.extrusionOk };

    return new BuilderElement({
        name: "a:path",
        attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
        children: options.commands.map(createPathCommand),
    });
};

// ─── Adjust Handle ──────────────────────────────────────────────────────────

/**
 * A position reference for adjust handles and connection sites.
 */
export interface AdjustHandlePosition {
    readonly x: string;
    readonly y: string;
}

/**
 * XY adjust handle (CT_XYAdjustHandle).
 *
 * Allows the user to drag a handle in X/Y directions.
 */
export interface XYAdjustHandle {
    readonly type: "xy";
    readonly gdRefX?: string;
    readonly minX?: string;
    readonly maxX?: string;
    readonly gdRefY?: string;
    readonly minY?: string;
    readonly maxY?: string;
    readonly position: AdjustHandlePosition;
}

/**
 * Polar adjust handle (CT_PolarAdjustHandle).
 *
 * Allows the user to drag a handle in radius/angle directions.
 */
export interface PolarAdjustHandle {
    readonly type: "polar";
    readonly gdRefR?: string;
    readonly minR?: string;
    readonly maxR?: string;
    readonly gdRefAng?: string;
    readonly minAng?: string;
    readonly maxAng?: string;
    readonly position: AdjustHandlePosition;
}

/** An adjust handle (XY or Polar). */
export type AdjustHandle = XYAdjustHandle | PolarAdjustHandle;

const createAdjustHandlePos = (position: AdjustHandlePosition): XmlComponent =>
    new BuilderElement<{ readonly x: string; readonly y: string }>({
        name: "a:pos",
        attributes: {
            x: { key: "x", value: position.x },
            y: { key: "y", value: position.y },
        },
    });

const createXYAdjustHandle = (handle: XYAdjustHandle): XmlComponent => {
    const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
    if (handle.gdRefX !== undefined) attrs.gdRefX = { key: "gdRefX", value: handle.gdRefX };
    if (handle.minX !== undefined) attrs.minX = { key: "minX", value: handle.minX };
    if (handle.maxX !== undefined) attrs.maxX = { key: "maxX", value: handle.maxX };
    if (handle.gdRefY !== undefined) attrs.gdRefY = { key: "gdRefY", value: handle.gdRefY };
    if (handle.minY !== undefined) attrs.minY = { key: "minY", value: handle.minY };
    if (handle.maxY !== undefined) attrs.maxY = { key: "maxY", value: handle.maxY };

    return new BuilderElement({
        name: "a:ahXY",
        attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
        children: [createAdjustHandlePos(handle.position)],
    });
};

const createPolarAdjustHandle = (handle: PolarAdjustHandle): XmlComponent => {
    const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
    if (handle.gdRefR !== undefined) attrs.gdRefR = { key: "gdRefR", value: handle.gdRefR };
    if (handle.minR !== undefined) attrs.minR = { key: "minR", value: handle.minR };
    if (handle.maxR !== undefined) attrs.maxR = { key: "maxR", value: handle.maxR };
    if (handle.gdRefAng !== undefined) attrs.gdRefAng = { key: "gdRefAng", value: handle.gdRefAng };
    if (handle.minAng !== undefined) attrs.minAng = { key: "minAng", value: handle.minAng };
    if (handle.maxAng !== undefined) attrs.maxAng = { key: "maxAng", value: handle.maxAng };

    return new BuilderElement({
        name: "a:ahPolar",
        attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
        children: [createAdjustHandlePos(handle.position)],
    });
};

// ─── Connection Site ────────────────────────────────────────────────────────

/**
 * Connection site (CT_ConnectionSite).
 *
 * Defines a point on the shape boundary where connectors can attach.
 */
export interface ConnectionSite {
    /** Angle (absolute value or guide name) */
    readonly ang: string;
    /** Position */
    readonly position: AdjustHandlePosition;
}

const createConnectionSite = (site: ConnectionSite): XmlComponent =>
    new BuilderElement<{ readonly ang: string }>({
        name: "a:cxn",
        attributes: { ang: { key: "ang", value: site.ang } },
        children: [createAdjustHandlePos(site.position)],
    });

// ─── Geometry Rectangle ─────────────────────────────────────────────────────

/**
 * Text insertion rectangle (CT_GeomRect).
 *
 * Defines the rectangle within the shape where text is placed.
 * Coordinates can be absolute values or references to geometry guide names.
 */
export interface GeomRect {
    readonly l: string;
    readonly t: string;
    readonly r: string;
    readonly b: string;
}

const createGeomRect = (rect: GeomRect): XmlComponent =>
    new BuilderElement<{
        readonly l: string;
        readonly t: string;
        readonly r: string;
        readonly b: string;
    }>({
        name: "a:rect",
        attributes: {
            l: { key: "l", value: rect.l },
            t: { key: "t", value: rect.t },
            r: { key: "r", value: rect.r },
            b: { key: "b", value: rect.b },
        },
    });

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
    readonly adjustmentValues?: readonly GeometryGuide[];
    /** Geometry guide formulas (a:gdLst) */
    readonly guides?: readonly GeometryGuide[];
    /** Adjust handles (a:ahLst) */
    readonly adjustHandles?: readonly AdjustHandle[];
    /** Connection sites (a:cxnLst) */
    readonly connectionSites?: readonly ConnectionSite[];
    /** Text insertion rectangle (a:rect) */
    readonly textRect?: GeomRect;
    /** Path definitions (a:pathLst) — required */
    readonly pathList: readonly PathOptions[];
}

/**
 * Creates a custom 2D geometry element (a:custGeom).
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
 * });
 * ```
 */
export const createCustomGeometry = (options: CustomGeometryOptions): XmlComponent => {
    const children: XmlComponent[] = [];

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
            new BuilderElement({
                name: "a:ahLst",
                children: options.adjustHandles.map((h) =>
                    h.type === "xy" ? createXYAdjustHandle(h) : createPolarAdjustHandle(h),
                ),
            }),
        );
    }

    // a:cxnLst
    if (options.connectionSites && options.connectionSites.length > 0) {
        children.push(
            new BuilderElement({
                name: "a:cxnLst",
                children: options.connectionSites.map(createConnectionSite),
            }),
        );
    }

    // a:rect
    if (options.textRect) {
        children.push(createGeomRect(options.textRect));
    }

    // a:pathLst (required)
    children.push(
        new BuilderElement({
            name: "a:pathLst",
            children: options.pathList.map(createPath),
        }),
    );

    return new BuilderElement({ name: "a:custGeom", children });
};
