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

import { xsdLightRigDirection } from "../../util/mappings";
import type { PositivePercentage } from "../../util/values";

/** Preset camera type (ST_PresetCameraType). */
export type CameraPreset =
  | "legacyObliqueTopLeft"
  | "legacyObliqueTop"
  | "legacyObliqueTopRight"
  | "legacyObliqueLeft"
  | "legacyObliqueFront"
  | "legacyObliqueRight"
  | "legacyObliqueBottomLeft"
  | "legacyObliqueBottom"
  | "legacyObliqueBottomRight"
  | "legacyPerspectiveTopLeft"
  | "legacyPerspectiveTop"
  | "legacyPerspectiveTopRight"
  | "legacyPerspectiveLeft"
  | "legacyPerspectiveFront"
  | "legacyPerspectiveRight"
  | "legacyPerspectiveBottomLeft"
  | "legacyPerspectiveBottom"
  | "legacyPerspectiveBottomRight"
  | "orthographicFront"
  | "isometricTopUp"
  | "isometricTopDown"
  | "isometricBottomUp"
  | "isometricBottomDown"
  | "isometricLeftUp"
  | "isometricLeftDown"
  | "isometricRightUp"
  | "isometricRightDown"
  | "isometricOffAxis1Left"
  | "isometricOffAxis1Right"
  | "isometricOffAxis1Top"
  | "isometricOffAxis2Left"
  | "isometricOffAxis2Right"
  | "isometricOffAxis2Top"
  | "isometricOffAxis3Left"
  | "isometricOffAxis3Right"
  | "isometricOffAxis3Bottom"
  | "isometricOffAxis4Left"
  | "isometricOffAxis4Right"
  | "isometricOffAxis4Bottom"
  | "obliqueTopLeft"
  | "obliqueTop"
  | "obliqueTopRight"
  | "obliqueLeft"
  | "obliqueRight"
  | "obliqueBottomLeft"
  | "obliqueBottom"
  | "obliqueBottomRight"
  | "perspectiveFront"
  | "perspectiveLeft"
  | "perspectiveRight"
  | "perspectiveAbove"
  | "perspectiveBelow"
  | "perspectiveAboveLeftFacing"
  | "perspectiveAboveRightFacing"
  | "perspectiveContrastingLeftFacing"
  | "perspectiveContrastingRightFacing"
  | "perspectiveHeroicLeftFacing"
  | "perspectiveHeroicRightFacing"
  | "perspectiveHeroicExtremeLeftFacing"
  | "perspectiveHeroicExtremeRightFacing"
  | "perspectiveRelaxed"
  | "perspectiveRelaxedModerately";

/** Light rig preset (ST_LightRigType). */
export type LightRigType =
  | "legacyFlat1"
  | "legacyFlat2"
  | "legacyFlat3"
  | "legacyFlat4"
  | "legacyNormal1"
  | "legacyNormal2"
  | "legacyNormal3"
  | "legacyNormal4"
  | "legacyHarsh1"
  | "legacyHarsh2"
  | "legacyHarsh3"
  | "legacyHarsh4"
  | "threePt"
  | "balanced"
  | "soft"
  | "harsh"
  | "flood"
  | "contrasting"
  | "morning"
  | "sunrise"
  | "sunset"
  | "chilly"
  | "freezing"
  | "flat"
  | "twoPt"
  | "glow"
  | "brightRoom";

/** Light direction (ST_LightRigDirection) — full words; XSD tokens topLeft→tl etc. */
export type LightRigDirection =
  | "topLeft"
  | "top"
  | "topRight"
  | "left"
  | "right"
  | "bottomLeft"
  | "bottom"
  | "bottomRight";

// ─── Sphere Coordinates ─────────────────────────────────────────────────────

/**
 * Sphere coordinates (CT_SphereCoords).
 *
 * Used for camera and light rig rotation. All angles are in degrees.
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
  /** Latitude angle in degrees (0–360). */
  lat: number;
  /** Longitude angle in degrees (0–360). */
  lon: number;
  /** Revolution angle in degrees (0–360). */
  rev: number;
}

const createSphereCoords = (coords: SphereCoords): string => {
  const lat = Math.round(coords.lat * 60000);
  const lon = Math.round(coords.lon * 60000);
  const rev = Math.round(coords.rev * 60000);
  return `<a:rot lat="${lat}" lon="${lon}" rev="${rev}"/>`;
};

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
  /** Preset camera type (ST_PresetCameraType) */
  preset: CameraPreset;
  /** Field of view angle in degrees (0–180). */
  fov?: number;
  /** Zoom percentage, e.g. "100%" (ST_PositivePercentage) */
  zoom?: PositivePercentage;
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
    {
      prst: options.preset,
      fov: options.fov !== undefined ? Math.round(options.fov * 60000) : undefined,
      zoom: options.zoom,
    },
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
  /** Light rig preset (ST_LightRigType) */
  rig: LightRigType;
  /** Light direction (ST_LightRigDirection, full words: "tl" → topLeft) */
  direction: LightRigDirection;
  /** Light rig rotation */
  rotation?: SphereCoords;
}

const createLightRig = (options: LightRigOptions): string => {
  const children: string[] = [];
  if (options.rotation) {
    children.push(createSphereCoords(options.rotation));
  }

  return element(
    "a:lightRig",
    { rig: options.rig, dir: xsdLightRigDirection.to(options.direction) },
    children,
  );
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
  /** Trailing a:extLst verbatim inner XML (unknown extensions). */
  ext?: string;
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
    ...(options.ext !== undefined ? [`<a:extLst>${options.ext}</a:extLst>`] : []),
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
 *   lightRig: { rig: "threePt", direction: "top" },
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
 *   lightRig: { rig: "threePt", direction: "top" },
 * });
 *
 * // Scene with rotated camera and backdrop
 * createScene3D({
 *   camera: {
 *     preset: "isometricTopUp",
 *     rotation: { lat: 0, lon: 0, rev: 90 },
 *   },
 *   lightRig: { rig: "balanced", direction: "topLeft" },
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
