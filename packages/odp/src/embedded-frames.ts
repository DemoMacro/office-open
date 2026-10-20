import { attributeString, childNamed, isExternalOdfUrl, xmlElement } from "@office-open/odf";
import type { ShapeOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import type { OdpEmbeddedObjectFrameOptions, OdpEmbeddedObjectOptions } from "./embedded-objects";
import { parseShape, toOdfLength } from "./shape";
import type { TextProperties } from "./text-run";

export type OdpEmbeddedObjectPool = Map<string, OdpEmbeddedObjectOptions>;

export function parseEmbeddedObjectFrame(
  frame: Element,
  objects: OdpEmbeddedObjectPool,
  binaries: Record<string, Uint8Array>,
  mediaTypes: Map<string, string | undefined>,
  nextShapeId: () => number,
  textStyles: Map<string, TextProperties>,
): { shape: ShapeOptions } | undefined {
  const reference = childNamed(frame, "draw:object");
  const sourceHref = attributeString(reference, "xlink:href");
  const external = sourceHref !== undefined && isExternalEmbeddedHref(sourceHref);
  const href = sourceHref
    ?.replace(/^\.\//, "")
    .replace(/^\//, "")
    .replace(/\/$/, "")
    .replace(/%(?![0-9A-Fa-f]{2})/g, "%25");
  const normalizedHref = href ? decodeURIComponent(href) : undefined;
  if (!href) return undefined;
  let object = objects.get(normalizedHref ?? "");
  if (!object) {
    if (!external) return undefined;
    object = { sourceUrl: sourceHref, members: [] };
    objects.set(normalizedHref ?? "", object);
  }
  const shapeId = nextShapeId();
  const frameOptions: OdpEmbeddedObjectFrameOptions = {
    shapeId,
    ...(attributeString(frame, "draw:style-name")
      ? { styleName: attributeString(frame, "draw:style-name") }
      : {}),
  };
  const replacement = childNamed(frame, "draw:image");
  const replacementHref = attributeString(replacement, "xlink:href")
    ?.replace(/^\.\//, "")
    .replace(/^\//, "");
  const replacementData = replacementHref ? binaries[replacementHref] : undefined;
  if (replacementHref && replacementData !== undefined) {
    frameOptions.replacementPath = replacementHref;
    const mediaType = mediaTypes.get(replacementHref);
    if (!object.members.some((member) => member.path === replacementHref))
      object.members.push({
        path: replacementHref,
        ...(mediaType === undefined ? {} : { mediaType }),
        data: replacementData,
      });
  }
  object.frames = [...(object.frames ?? []), frameOptions];
  return {
    shape: parseShape(frame, textStyles, shapeId),
  };
}

function isExternalEmbeddedHref(href: string): boolean {
  return isExternalOdfUrl(href) || href.startsWith("///");
}

export function embeddedObjectXml(
  shape: ShapeOptions,
  object: OdpEmbeddedObjectOptions,
  frame: OdpEmbeddedObjectFrameOptions,
): string {
  const href = object.path ? `./${object.path}` : object.sourceUrl;
  if (!href) throw new Error("Embedded ODF object has no package path or source URL");
  return xmlElement(
    "draw:frame",
    {
      "draw:name": shape.name,
      "draw:style-name": frame.styleName,
      "svg:x": toOdfLength(shape.x),
      "svg:y": toOdfLength(shape.y),
      "svg:width": toOdfLength(shape.width),
      "svg:height": toOdfLength(shape.height),
    },
    [
      xmlElement("draw:object", {
        "xlink:href": href,
        "xlink:type": "simple",
        "xlink:show": "embed",
        "xlink:actuate": "onLoad",
      }),
      ...(frame.replacementPath
        ? [
            xmlElement("draw:image", {
              "xlink:href": `./${frame.replacementPath}`,
              "xlink:type": "simple",
              "xlink:show": "embed",
              "xlink:actuate": "onLoad",
            }),
          ]
        : []),
    ],
  );
}
