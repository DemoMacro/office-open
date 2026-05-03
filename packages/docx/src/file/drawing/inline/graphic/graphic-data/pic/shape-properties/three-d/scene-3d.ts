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

export { createScene3D } from "@office-open/core/drawingml";
export type {
    Scene3DOptions,
    CameraOptions,
    LightRigOptions,
    BackdropOptions,
    SphereCoords,
    Point3D,
    Vector3D,
} from "@office-open/core/drawingml";
