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
import { element } from "@office-open/xml";

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
  lat: number;
  /** Longitude angle (0 to 21600000) */
  lon: number;
  /** Revolution angle (0 to 21600000) */
  rev: number;
}

const createSphereCoords = (coords: SphereCoords): string =>
  `<a:rot lat="${coords.lat}" lon="${coords.lon}" rev="${coords.rev}"/>`;

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
  preset: string;
  /** Field of view angle (0 to 10800000, in 1/60000 degree units, max 180°) */
  fov?: number;
  /** Zoom percentage (e.g., "100%") */
  zoom?: string;
  /** Camera rotation */
  rotation?: SphereCoords;
}

const createCamera = (options: CameraOptions): string => {
  const children: string[] = [];
  if (options.rotation) {
    children.push(createSphereCoords(options.rotation));
  }

  return element(
    "a:camera",
    { prst: options.preset, fov: options.fov, zoom: options.zoom },
    children,
  );
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
  rig: string;
  /** Light direction (e.g., "tl", "t", "tr", "l", "r", "bl", "b", "br") */
  direction: string;
  /** Light rig rotation */
  rotation?: SphereCoords;
}

const createLightRig = (options: LightRigOptions): string => {
  const children: string[] = [];
  if (options.rotation) {
    children.push(createSphereCoords(options.rotation));
  }

  return element("a:lightRig", { rig: options.rig, dir: options.direction }, children);
};

// ─── Backdrop ───────────────────────────────────────────────────────────────

/**
 * 3D point (CT_Point3D).
 */
export interface Point3D {
  x: number;
  y: number;
  z: number;
}

/**
 * 3D vector (CT_Vector3D).
 */
export interface Vector3D {
  dx: number;
  dy: number;
  dz: number;
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
  anchor: Point3D;
  /** Normal vector */
  normal: Vector3D;
  /** Up vector */
  up: Vector3D;
}

const createPoint3D = (name: string, point: Point3D): string =>
  `<${name} x="${point.x}" y="${point.y}" z="${point.z}"/>`;

const createVector3D = (name: string, vector: Vector3D): string =>
  `<${name} dx="${vector.dx}" dy="${vector.dy}" dz="${vector.dz}"/>`;

const createBackdrop = (options: BackdropOptions): string =>
  element("a:backdrop", undefined, [
    createPoint3D("a:anchor", options.anchor),
    createVector3D("a:norm", options.normal),
    createVector3D("a:up", options.up),
  ]);

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
  camera: CameraOptions;
  /** Light rig settings (required) */
  lightRig: LightRigOptions;
  /** Backdrop settings (optional) */
  backdrop?: BackdropOptions;
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
export const createScene3D = (options: Scene3DOptions): string => {
  const children: string[] = [createCamera(options.camera), createLightRig(options.lightRig)];

  if (options.backdrop) {
    children.push(createBackdrop(options.backdrop));
  }

  return element("a:scene3d", undefined, children);
};
