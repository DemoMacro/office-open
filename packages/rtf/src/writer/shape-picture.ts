import type { PictureOptions, ShapeOptions } from "@office-open/docx";

import { assertAllowed, reject } from "./context";
import { control, rtfText } from "./escape";

const EMU_PER_TWIP = 635;

function twips(value: number, path: string): number {
  if (!Number.isFinite(value) || value / EMU_PER_TWIP !== Math.round(value / EMU_PER_TWIP)) {
    reject("picture", path, "extent", "RTF picture extents must resolve to whole twips");
  }
  return Math.round(value / EMU_PER_TWIP);
}

export function writePicture(picture: PictureOptions, path: string): string {
  assertAllowed(picture, ["type", "data", "transformation"], "picture", path);
  assertAllowed(picture.transformation, ["width", "height"], "picture", `${path}.transformation`);
  if (
    typeof picture.transformation.width !== "number" ||
    typeof picture.transformation.height !== "number"
  ) {
    reject("picture", `${path}.transformation`, "width", "RTF requires numeric EMU extents");
  }
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
  const width = twips(picture.transformation.width, `${path}.transformation.width`);
  const height = twips(picture.transformation.height, `${path}.transformation.height`);
  if (!(picture.data instanceof Uint8Array)) {
    reject("picture", `${path}.data`, "data", "RTF pictures require binary data");
  }
  let data = "";
  for (const byte of picture.data) data += byte.toString(16).padStart(2, "0");
  return `{\\pict${typeControl}${control("picw", width)}${control("pich", height)} ${data}}`;
}

export function writeShape(shape: ShapeOptions, path: string): string {
  assertAllowed(shape, ["children", "transformation", "floating"], "shape", path);
  assertAllowed(shape.transformation, ["width", "height"], "shape", `${path}.transformation`);
  if (
    typeof shape.transformation.width !== "number" ||
    typeof shape.transformation.height !== "number"
  ) {
    reject("shape", `${path}.transformation`, "width", "RTF requires numeric EMU extents");
  }
  const floating = shape.floating;
  if (floating) {
    assertAllowed(
      floating,
      ["horizontalPosition", "verticalPosition"],
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
    left === undefined ? "" : control("shpleft", Math.round(left / EMU_PER_TWIP)),
    top === undefined ? "" : control("shptop", Math.round(top / EMU_PER_TWIP)),
    control("shpwidth", Math.round(shape.transformation.width / EMU_PER_TWIP)),
    control("shpheight", Math.round(shape.transformation.height / EMU_PER_TWIP)),
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
    return `{\\object${className ? `{\\objclass ${className}}` : ""}{\\objdata${control("bin", object.embed!.data.length)} ${binary}}}`;
  }
  return undefined;
}
