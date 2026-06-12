/**
 * 3D descriptor for DrawingML shapes.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import { xsdMaterialType } from "../../util/mappings";
import { solidFillDesc } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import type { BevelOptions } from "./bevel";
import type {
  Scene3DOptions,
  CameraOptions,
  LightRigOptions,
  SphereCoords,
  BackdropOptions,
  Point3D,
  Vector3D,
} from "./scene-3d";
import type { Shape3DOptions } from "./shape-3d";

// ── SphereCoords helper ──

function stringifySphereCoords(coords: SphereCoords): string {
  return `<a:rot lat="${coords.lat}" lon="${coords.lon}" rev="${coords.rev}"/>`;
}

function readSphereCoords(el: XmlElement): SphereCoords | undefined {
  const lat = el.attributes?.["lat"];
  const lon = el.attributes?.["lon"];
  const rev = el.attributes?.["rev"];
  if (lat === undefined || lon === undefined || rev === undefined) return undefined;
  return { lat: Number(lat), lon: Number(lon), rev: Number(rev) };
}

// ── Bevel descriptor (a:bevelT / a:bevelB) ──

export const bevelDesc: CustomDescriptor<BevelOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    if (opts.w !== undefined) attrParts.push(`w="${opts.w}"`);
    if (opts.h !== undefined) attrParts.push(`h="${opts.h}"`);
    if (opts.prst !== undefined) attrParts.push(`prst="${escapeXml(opts.prst)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:bevel${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: Partial<BevelOptions> = {};
    if (el.attributes?.["w"] !== undefined) result.w = Number(el.attributes["w"]);
    if (el.attributes?.["h"] !== undefined) result.h = Number(el.attributes["h"]);
    if (el.attributes?.["prst"] !== undefined)
      result.prst = String(el.attributes["prst"]) as BevelOptions["prst"];
    return result as BevelOptions;
  },
};

// ── Shape3D descriptor (a:sp3d) ──

export const shape3DDesc: CustomDescriptor<Shape3DOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];

    // Children
    if (opts.bevelT) {
      const bevelXml = stringify(bevelDesc, opts.bevelT, ctx);
      if (bevelXml) parts.push(bevelXml.replace("<a:bevel", "<a:bevelT"));
    }
    if (opts.bevelB) {
      const bevelXml = stringify(bevelDesc, opts.bevelB, ctx);
      if (bevelXml) parts.push(bevelXml.replace("<a:bevel", "<a:bevelB"));
    }
    if (opts.extrusionColor) {
      const colorXml = stringify(solidFillDesc, opts.extrusionColor, ctx);
      if (colorXml) parts.push(`<a:extrusionClr>${colorXml}</a:extrusionClr>`);
    }
    if (opts.contourColor) {
      const colorXml = stringify(solidFillDesc, opts.contourColor, ctx);
      if (colorXml) parts.push(`<a:contourClr>${colorXml}</a:contourClr>`);
    }

    // Attributes
    const attrParts: string[] = [];
    if (opts.z !== undefined) attrParts.push(`z="${opts.z}"`);
    if (opts.extrusionH !== undefined) attrParts.push(`extrusionH="${opts.extrusionH}"`);
    if (opts.contourW !== undefined) attrParts.push(`contourW="${opts.contourW}"`);
    if (opts.prstMaterial !== undefined) {
      attrParts.push(`prstMaterial="${escapeXml(xsdMaterialType.to(opts.prstMaterial))}"`);
    }
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    const content = parts.join("");
    if (!attrStr && !content) return undefined;
    if (!content) return `<a:sp3d${attrStr}/>`;
    return `<a:sp3d${attrStr}>${content}</a:sp3d>`;
  },
  parse(el, ctx) {
    const result: Partial<Shape3DOptions> = {};

    // Attributes
    if (el.attributes?.["z"] !== undefined) result.z = Number(el.attributes["z"]);
    if (el.attributes?.["extrusionH"] !== undefined)
      result.extrusionH = Number(el.attributes["extrusionH"]);
    if (el.attributes?.["contourW"] !== undefined)
      result.contourW = Number(el.attributes["contourW"]);
    if (el.attributes?.["prstMaterial"] !== undefined) {
      result.prstMaterial = xsdMaterialType.from(
        String(el.attributes["prstMaterial"]),
      ) as Shape3DOptions["prstMaterial"];
    }

    // Children
    const bevelT = findChild(el, "a:bevelT");
    if (bevelT) result.bevelT = parse(bevelDesc, bevelT, ctx) as BevelOptions;

    const bevelB = findChild(el, "a:bevelB");
    if (bevelB) result.bevelB = parse(bevelDesc, bevelB, ctx) as BevelOptions;

    const extrusionClr = findChild(el, "a:extrusionClr");
    if (extrusionClr) {
      const solidFill = findChild(extrusionClr, "a:solidFill");
      if (solidFill)
        result.extrusionColor = parse(solidFillDesc, solidFill, ctx) as SolidFillOptions;
    }

    const contourClr = findChild(el, "a:contourClr");
    if (contourClr) {
      const solidFill = findChild(contourClr, "a:solidFill");
      if (solidFill) result.contourColor = parse(solidFillDesc, solidFill, ctx) as SolidFillOptions;
    }

    return result as Shape3DOptions;
  },
};

// ── Camera descriptor ──

const cameraDesc: CustomDescriptor<CameraOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    attrParts.push(`prst="${escapeXml(opts.preset)}"`);
    if (opts.fov !== undefined) attrParts.push(`fov="${opts.fov}"`);
    if (opts.zoom !== undefined) attrParts.push(`zoom="${escapeXml(opts.zoom)}"`);
    const attrStr = " " + attrParts.join(" ");

    const parts: string[] = [];
    if (opts.rotation) parts.push(stringifySphereCoords(opts.rotation));

    const content = parts.join("");
    if (!content) return `<a:camera${attrStr}/>`;
    return `<a:camera${attrStr}>${content}</a:camera>`;
  },
  parse(el, _ctx) {
    const result: Partial<CameraOptions> = {};
    if (el.attributes?.["prst"] !== undefined) result.preset = String(el.attributes["prst"]);
    if (el.attributes?.["fov"] !== undefined) result.fov = Number(el.attributes["fov"]);
    if (el.attributes?.["zoom"] !== undefined) result.zoom = String(el.attributes["zoom"]);
    const rot = findChild(el, "a:rot");
    if (rot) result.rotation = readSphereCoords(rot);
    return result as CameraOptions;
  },
};

// ── LightRig descriptor ──

const lightRigDesc: CustomDescriptor<LightRigOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    attrParts.push(`rig="${escapeXml(opts.rig)}"`);
    attrParts.push(`dir="${escapeXml(opts.direction)}"`);
    const attrStr = " " + attrParts.join(" ");

    const parts: string[] = [];
    if (opts.rotation) parts.push(stringifySphereCoords(opts.rotation));

    const content = parts.join("");
    if (!content) return `<a:lightRig${attrStr}/>`;
    return `<a:lightRig${attrStr}>${content}</a:lightRig>`;
  },
  parse(el, _ctx) {
    const result: Partial<LightRigOptions> = {};
    if (el.attributes?.["rig"] !== undefined) result.rig = String(el.attributes["rig"]);
    if (el.attributes?.["dir"] !== undefined) result.direction = String(el.attributes["dir"]);
    const rot = findChild(el, "a:rot");
    if (rot) result.rotation = readSphereCoords(rot);
    return result as LightRigOptions;
  },
};

// ── Scene3D descriptor (a:scene3d) ──

export const scene3DDesc: CustomDescriptor<Scene3DOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];

    // Camera (required)
    const cameraXml = stringify(cameraDesc, opts.camera, ctx);
    if (cameraXml) parts.push(cameraXml);

    // Light rig (required)
    const lightRigXml = stringify(lightRigDesc, opts.lightRig, ctx);
    if (lightRigXml) parts.push(lightRigXml);

    // Backdrop (optional)
    if (opts.backdrop) {
      parts.push(stringifyBackdrop(opts.backdrop));
    }

    const content = parts.join("");
    if (!content) return undefined;
    return `<a:scene3d>${content}</a:scene3d>`;
  },
  parse(el, ctx) {
    const result: Partial<Scene3DOptions> = {};

    const camera = findChild(el, "a:camera");
    if (camera) result.camera = parse(cameraDesc, camera, ctx) as CameraOptions;

    const lightRig = findChild(el, "a:lightRig");
    if (lightRig) result.lightRig = parse(lightRigDesc, lightRig, ctx) as LightRigOptions;

    const backdrop = findChild(el, "a:backdrop");
    if (backdrop) result.backdrop = readBackdrop(backdrop);

    return result as Scene3DOptions;
  },
};

// ── Backdrop helpers ──

function stringifyPoint3D(name: string, point: Point3D): string {
  return `<${name} x="${point.x}" y="${point.y}" z="${point.z}"/>`;
}

function stringifyVector3D(name: string, vector: Vector3D): string {
  return `<${name} dx="${vector.dx}" dy="${vector.dy}" dz="${vector.dz}"/>`;
}

function stringifyBackdrop(opts: BackdropOptions): string {
  const content =
    stringifyPoint3D("a:anchor", opts.anchor) +
    stringifyVector3D("a:norm", opts.normal) +
    stringifyVector3D("a:up", opts.up);
  return `<a:backdrop>${content}</a:backdrop>`;
}

function readPoint3D(el: XmlElement): Point3D | undefined {
  const x = el.attributes?.["x"];
  const y = el.attributes?.["y"];
  const z = el.attributes?.["z"];
  if (x === undefined || y === undefined || z === undefined) return undefined;
  return { x: Number(x), y: Number(y), z: Number(z) };
}

function readVector3D(el: XmlElement): Vector3D | undefined {
  const dx = el.attributes?.["dx"];
  const dy = el.attributes?.["dy"];
  const dz = el.attributes?.["dz"];
  if (dx === undefined || dy === undefined || dz === undefined) return undefined;
  return { dx: Number(dx), dy: Number(dy), dz: Number(dz) };
}

function readBackdrop(el: XmlElement): BackdropOptions | undefined {
  const anchor = findChild(el, "a:anchor");
  const norm = findChild(el, "a:norm");
  const up = findChild(el, "a:up");
  if (!anchor || !norm || !up) return undefined;
  const anchorPt = readPoint3D(anchor);
  const normVec = readVector3D(norm);
  const upVec = readVector3D(up);
  if (!anchorPt || !normVec || !upVec) return undefined;
  return { anchor: anchorPt, normal: normVec, up: upVec };
}
