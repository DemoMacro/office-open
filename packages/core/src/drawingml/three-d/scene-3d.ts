/**
 * 3D scene module for DrawingML shapes.
 *
 * Provides CT_Scene3D — the 3D scene properties including camera, light rig,
 * and optional backdrop that define how a 3D shape is rendered.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Scene3D
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";

// ─── Sphere Coordinates ─────────────────────────────────────────────────────

/**
 * Sphere coordinates (CT_SphereCoords).
 *
 * Used for camera and light rig rotation. All angles are in 1/60000 of a degree.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SphereCoords">
 *   <xsd:attribute name="lat" type="ST_PositiveFixedAngle" use="required"/>
 *   <xsd:attribute name="lon" type="ST_PositiveFixedAngle" use="required"/>
 *   <xsd:attribute name="rev" type="ST_PositiveFixedAngle" use="required"/>
 * </xsd:complexType>
 * ```
 */
export interface SphereCoords {
    /** Latitude angle (0 to 21600000, in 1/60000 degree units) */
    readonly lat: number;
    /** Longitude angle (0 to 21600000) */
    readonly lon: number;
    /** Revolution angle (0 to 21600000) */
    readonly rev: number;
}

const createSphereCoords = (coords: SphereCoords): XmlComponent =>
    new BuilderElement<{ readonly lat: number; readonly lon: number; readonly rev: number }>({
        name: "a:rot",
        attributes: {
            lat: { key: "lat", value: coords.lat },
            lon: { key: "lon", value: coords.lon },
            rev: { key: "rev", value: coords.rev },
        },
    });

// ─── Camera ─────────────────────────────────────────────────────────────────

/**
 * Camera options (CT_Camera).
 *
 * Defines the camera type and optional rotation for the 3D scene.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Camera">
 *   <xsd:sequence>
 *     <xsd:element name="rot" type="CT_SphereCoords" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="prst" type="ST_PresetCameraType" use="required"/>
 *   <xsd:attribute name="fov" type="ST_FOVAngle" use="optional"/>
 *   <xsd:attribute name="zoom" type="ST_PositivePercentage" use="optional" default="100%"/>
 * </xsd:complexType>
 * ```
 */
export interface CameraOptions {
    /** Preset camera type (e.g., "perspectiveFront", "isometricTopUp") */
    readonly preset: string;
    /** Field of view angle (0 to 10800000, in 1/60000 degree units, max 180°) */
    readonly fov?: number;
    /** Zoom percentage (e.g., "100%") */
    readonly zoom?: string;
    /** Camera rotation */
    readonly rotation?: SphereCoords;
}

const createCamera = (options: CameraOptions): XmlComponent => {
    const children: XmlComponent[] = [];
    if (options.rotation) {
        children.push(createSphereCoords(options.rotation));
    }

    const attrs: Record<string, { readonly key: string; readonly value: string | number }> = {
        prst: { key: "prst", value: options.preset },
    };
    if (options.fov !== undefined) attrs.fov = { key: "fov", value: options.fov };
    if (options.zoom !== undefined) attrs.zoom = { key: "zoom", value: options.zoom };

    return new BuilderElement({
        name: "a:camera",
        attributes: attrs,
        children: children.length > 0 ? children : undefined,
    });
};

// ─── Light Rig ──────────────────────────────────────────────────────────────

/**
 * Light rig options (CT_LightRig).
 *
 * Defines the lighting setup for the 3D scene.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LightRig">
 *   <xsd:sequence>
 *     <xsd:element name="rot" type="CT_SphereCoords" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rig" type="ST_LightRigType" use="required"/>
 *   <xsd:attribute name="dir" type="ST_LightRigDirection" use="required"/>
 * </xsd:complexType>
 * ```
 */
export interface LightRigOptions {
    /** Light rig type (e.g., "threePt", "balanced", "soft") */
    readonly rig: string;
    /** Light direction (e.g., "tl", "t", "tr", "l", "r", "bl", "b", "br") */
    readonly direction: string;
    /** Light rig rotation */
    readonly rotation?: SphereCoords;
}

const createLightRig = (options: LightRigOptions): XmlComponent => {
    const children: XmlComponent[] = [];
    if (options.rotation) {
        children.push(createSphereCoords(options.rotation));
    }

    return new BuilderElement<{ readonly rig: string; readonly dir: string }>({
        name: "a:lightRig",
        attributes: {
            rig: { key: "rig", value: options.rig },
            dir: { key: "dir", value: options.direction },
        },
        children: children.length > 0 ? children : undefined,
    });
};

// ─── Backdrop ───────────────────────────────────────────────────────────────

/**
 * 3D point (CT_Point3D).
 */
export interface Point3D {
    readonly x: number;
    readonly y: number;
    readonly z: number;
}

/**
 * 3D vector (CT_Vector3D).
 */
export interface Vector3D {
    readonly dx: number;
    readonly dy: number;
    readonly dz: number;
}

/**
 * Backdrop options (CT_Backdrop).
 *
 * Defines the backdrop plane for the 3D scene.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Backdrop">
 *   <xsd:sequence>
 *     <xsd:element name="anchor" type="CT_Point3D" use="required"/>
 *     <xsd:element name="norm" type="CT_Vector3D" use="required"/>
 *     <xsd:element name="up" type="CT_Vector3D" use="required"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export interface BackdropOptions {
    /** Anchor point */
    readonly anchor: Point3D;
    /** Normal vector */
    readonly normal: Vector3D;
    /** Up vector */
    readonly up: Vector3D;
}

const createPoint3D = (name: string, point: Point3D): XmlComponent =>
    new BuilderElement<{ readonly x: number; readonly y: number; readonly z: number }>({
        name,
        attributes: {
            x: { key: "x", value: point.x },
            y: { key: "y", value: point.y },
            z: { key: "z", value: point.z },
        },
    });

const createVector3D = (name: string, vector: Vector3D): XmlComponent =>
    new BuilderElement<{ readonly dx: number; readonly dy: number; readonly dz: number }>({
        name,
        attributes: {
            dx: { key: "dx", value: vector.dx },
            dy: { key: "dy", value: vector.dy },
            dz: { key: "dz", value: vector.dz },
        },
    });

const createBackdrop = (options: BackdropOptions): XmlComponent =>
    new BuilderElement({
        name: "a:backdrop",
        children: [
            createPoint3D("a:anchor", options.anchor),
            createVector3D("a:norm", options.normal),
            createVector3D("a:up", options.up),
        ],
    });

// ─── Scene 3D ───────────────────────────────────────────────────────────────

/**
 * Options for a 3D scene (CT_Scene3D).
 *
 * Both `camera` and `lightRig` are required. `backdrop` is optional.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Scene3D">
 *   <xsd:sequence>
 *     <xsd:element name="camera" type="CT_Camera" minOccurs="1"/>
 *     <xsd:element name="lightRig" type="CT_LightRig" minOccurs="1"/>
 *     <xsd:element name="backdrop" type="CT_Backdrop" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * createScene3D({
 *   camera: { preset: "perspectiveFront" },
 *   lightRig: { rig: "threePt", direction: "t" },
 * });
 * ```
 */
export interface Scene3DOptions {
    /** Camera settings (required) */
    readonly camera: CameraOptions;
    /** Light rig settings (required) */
    readonly lightRig: LightRigOptions;
    /** Backdrop settings (optional) */
    readonly backdrop?: BackdropOptions;
}

/**
 * Creates a 3D scene element (a:scene3d).
 *
 * @example
 * ```typescript
 * // Simple scene with default camera and lighting
 * createScene3D({
 *   camera: { preset: "perspectiveFront" },
 *   lightRig: { rig: "threePt", direction: "t" },
 * });
 *
 * // Scene with rotated camera and backdrop
 * createScene3D({
 *   camera: {
 *     preset: "isometricTopUp",
 *     rotation: { lat: 0, lon: 0, rev: 5400000 },
 *   },
 *   lightRig: { rig: "balanced", direction: "tl" },
 *   backdrop: {
 *     anchor: { x: 0, y: 0, z: 0 },
 *     normal: { dx: 0, dy: 0, dz: 1 },
 *     up: { dx: 0, dy: 1, dz: 0 },
 *   },
 * });
 * ```
 */
export const createScene3D = (options: Scene3DOptions): XmlComponent => {
    const children: XmlComponent[] = [
        createCamera(options.camera),
        createLightRig(options.lightRig),
    ];

    if (options.backdrop) {
        children.push(createBackdrop(options.backdrop));
    }

    return new BuilderElement({ name: "a:scene3d", children });
};
