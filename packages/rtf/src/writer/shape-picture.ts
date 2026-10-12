import { convertToEmu } from "@office-open/core";
import type { MediaTransformation, PictureOptions, ShapeOptions } from "@office-open/docx";

import { assertAllowed, reject } from "./context";
import { control, rtfText } from "./escape";

const EMU_PER_TWIP = 635;

type RtfTransformation = MediaTransformation;

function transformationExtents(
  transformation: RtfTransformation,
  owner: "picture" | "shape",
  path: string,
): { width: number; height: number } {
  assertAllowed(transformation, ["width", "height"], owner, `${path}.transformation`);
  if (transformation.offset?.x || transformation.offset?.y) {
    reject(owner, `${path}.transformation.offset`, "offset", "RTF cannot express drawing offsets");
  }
  return {
    width: convertToEmu(transformation.width),
    height: convertToEmu(transformation.height),
  };
}

function twips(value: number, path: string): number {
  if (!Number.isFinite(value) || value / EMU_PER_TWIP !== Math.round(value / EMU_PER_TWIP)) {
    reject("picture", path, "extent", "RTF picture extents must resolve to whole twips");
  }
  return Math.round(value / EMU_PER_TWIP);
}

function wrapControls(wrap: NonNullable<ShapeOptions["floating"]>["wrap"], path: string): string {
  if (!wrap) return "";
  const type =
    wrap.type === "topAndBottom"
      ? 1
      : wrap.type === "square"
        ? 2
        : wrap.type === "none"
          ? 3
          : wrap.type === "tight"
            ? 4
            : wrap.type === "through"
              ? 5
              : reject("shape", `${path}.floating.wrap`, wrap.type, "unsupported RTF wrap type");
  let side = "";
  if (wrap.side !== undefined) {
    if (wrap.type !== "square" && wrap.type !== "tight" && wrap.type !== "through") {
      reject("shape", `${path}.floating.wrap.side`, wrap.side, "unsupported wrap side");
    }
    side = control(
      "shpwrk",
      wrap.side === "left" ? 1 : wrap.side === "right" ? 2 : wrap.side === "largest" ? 3 : 0,
    );
  }
  return `${control("shpwr", type)}${side}`;
}

function shapeOriginControls(shape: ShapeOptions, path: string): string {
  const floating = shape.floating;
  if (!floating) return "";
  const horizontal = floating.horizontalPosition?.relative ?? "column";
  const vertical = floating.verticalPosition?.relative ?? "paragraph";
  const horizontalControl =
    horizontal === "column"
      ? ""
      : horizontal === "margin"
        ? control("shpbxmargin")
        : horizontal === "page"
          ? control("shpbxpage")
          : reject(
              "shape",
              `${path}.floating.horizontalPosition.relative`,
              horizontal,
              "unsupported RTF horizontal origin",
            );
  const verticalControl =
    vertical === "paragraph"
      ? ""
      : vertical === "margin"
        ? control("shpbymargin")
        : vertical === "page"
          ? control("shpbypage")
          : reject(
              "shape",
              `${path}.floating.verticalPosition.relative`,
              vertical,
              "unsupported RTF vertical origin",
            );
  return `${horizontalControl}${verticalControl}`;
}

export function writePicture(picture: PictureOptions, path: string): string {
  assertAllowed(picture, ["type", "data", "transformation"], "picture", path);
  const typeControl =
    picture.type === "png"
      ? control("pngblip")
      : picture.type === "jpg"
        ? control("jpegblip")
        : picture.type === "emf"
          ? control("emfblip")
          : picture.type === "wmf"
            ? `${control("wmetafile", 8)}`
            : reject("picture", path, picture.type, "unsupported RTF picture type");
  const extents = transformationExtents(picture.transformation, "picture", path);
  const width = twips(extents.width, `${path}.transformation.width`);
  const height = twips(extents.height, `${path}.transformation.height`);
  if (!(picture.data instanceof Uint8Array)) {
    reject("picture", `${path}.data`, "data", "RTF pictures require binary data");
  }
  let data = "";
  for (const byte of picture.data) data += byte.toString(16).padStart(2, "0");
  return `{\\pict${typeControl}${control("picw", width)}${control("pich", height)} ${data}}`;
}

export function writeShape(shape: ShapeOptions, path: string): string {
  assertAllowed(shape, ["children", "transformation", "floating"], "shape", path);
  const extents = transformationExtents(shape.transformation, "shape", path);
  const floating = shape.floating;
  if (floating) {
    assertAllowed(
      floating,
      ["horizontalPosition", "verticalPosition", "behindDocument", "lockAnchor", "wrap", "zIndex"],
      "shape",
      `${path}.floating`,
    );
  }
  const left = floating?.horizontalPosition?.offset;
  const top = floating?.verticalPosition?.offset;
  if (left !== undefined && typeof left !== "number") {
    reject(
      "shape",
      `${path}.floating.horizontalPosition`,
      "offset",
      "RTF shape offsets require numeric EMU",
    );
  }
  if (top !== undefined && typeof top !== "number") {
    reject(
      "shape",
      `${path}.floating.verticalPosition`,
      "offset",
      "RTF shape offsets require numeric EMU",
    );
  }
  const text = shape.children
    .filter((child): child is string => typeof child === "string")
    .join("");
  if (shape.children.some((child) => typeof child !== "string")) {
    reject("shape", `${path}.children`, "children", "RTF shapes support text children only");
  }
  return [
    "{\\shpinst",
    shapeOriginControls(shape, path),
    left === undefined ? "" : control("shpleft", Math.round(left / EMU_PER_TWIP)),
    top === undefined ? "" : control("shptop", Math.round(top / EMU_PER_TWIP)),
    control("shpwidth", Math.round(extents.width / EMU_PER_TWIP)),
    control("shpheight", Math.round(extents.height / EMU_PER_TWIP)),
    floating?.behindDocument ? control("shpfblwtxt", 1) : "",
    floating?.lockAnchor ? control("shplockanchor") : "",
    floating?.zIndex === undefined ? "" : control("shpz", floating.zIndex),
    wrapControls(floating?.wrap, path),
    text ? `${control("shptxt")}${rtfText(text)}` : "",
    "}",
  ].join("");
}

export function writeDrawingChild(child: object, path: string): string | undefined {
  if ("picture" in child) {
    const picture = child.picture;
    if (typeof picture !== "object" || picture === null) return undefined;
    return writePicture(picture as PictureOptions, `${path}.picture`);
  }
  if ("wpsShape" in child) {
    const shape = child.wpsShape;
    if (typeof shape !== "object" || shape === null) return undefined;
    return writeShape(shape as ShapeOptions, `${path}.wpsShape`);
  }
  if ("object" in child) {
    const object = child.object as { embed?: { data?: Uint8Array; progId?: string } };
    if (
      typeof object !== "object" ||
      object === null ||
      !(object.embed?.data instanceof Uint8Array)
    )
      reject("object", `${path}.object`, "embed", "RTF objects require embedded binary data");
    let binary = "";
    for (const byte of object.embed!.data) binary += String.fromCharCode(byte);
    const className = object.embed!.progId;
    return `{\\object${className ? `{\\objclass ${className}}` : ""}{\\objdata${control("bin", object.embed!.data.length)}${binary}}}`;
  }
  return undefined;
}
