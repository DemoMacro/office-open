import { escapeText, xmlElement } from "@office-open/ocf";
import { parse, textOf, type Element } from "@office-open/xml";

import { OdfSchemaError } from "./error";

export const DRAWING_VOCABULARY_ELEMENTS = [
  "dr3d:cube",
  "dr3d:extrude",
  "dr3d:light",
  "dr3d:rotate",
  "dr3d:scene",
  "dr3d:sphere",
  "svg:definition-src",
  "svg:desc",
  "svg:font-face-format",
  "svg:font-face-name",
  "svg:font-face-src",
  "svg:font-face-uri",
  "svg:linearGradient",
  "svg:radialGradient",
  "svg:stop",
  "svg:title",
] as const;

export const DRAWING_VOCABULARY_ATTRIBUTES = [
  "draw:class-names",
  "draw:caption-id",
  "draw:display-name",
  "draw:id",
  "draw:layer",
  "draw:name",
  "draw:style-name",
  "draw:z-index",
  "dr3d:ambient-color",
  "dr3d:back-scale",
  "dr3d:backface-culling",
  "dr3d:center",
  "dr3d:close-back",
  "dr3d:close-front",
  "dr3d:depth",
  "dr3d:diffuse-color",
  "dr3d:direction",
  "dr3d:edge-rounding",
  "dr3d:edge-rounding-mode",
  "dr3d:emissive-color",
  "dr3d:enabled",
  "dr3d:focal-length",
  "dr3d:horizontal-segments",
  "dr3d:lighting-mode",
  "dr3d:max-edge",
  "dr3d:min-edge",
  "dr3d:normals-direction",
  "dr3d:normals-kind",
  "dr3d:projection",
  "dr3d:shadow-slant",
  "dr3d:shade-mode",
  "dr3d:shininess",
  "dr3d:shadow",
  "dr3d:specular",
  "dr3d:specular-color",
  "dr3d:texture-filter",
  "dr3d:texture-generation-mode-x",
  "dr3d:texture-generation-mode-y",
  "dr3d:texture-kind",
  "dr3d:texture-mode",
  "dr3d:transform",
  "dr3d:vertical-segments",
  "dr3d:vpn",
  "dr3d:vrp",
  "dr3d:vup",
  "presentation:class-names",
  "presentation:style-name",
  "svg:accent-height",
  "svg:alphabetic",
  "svg:ascent",
  "svg:bbox",
  "svg:cap-height",
  "svg:cx",
  "svg:cy",
  "svg:d",
  "svg:descent",
  "svg:fill-rule",
  "svg:font-family",
  "svg:font-size",
  "svg:font-stretch",
  "svg:font-style",
  "svg:font-variant",
  "svg:font-weight",
  "svg:fx",
  "svg:fy",
  "svg:gradientTransform",
  "svg:gradientUnits",
  "svg:hanging",
  "svg:height",
  "svg:ideographic",
  "svg:mathematical",
  "svg:name",
  "svg:offset",
  "svg:origin",
  "svg:overline-position",
  "svg:overline-thickness",
  "svg:panose-1",
  "svg:path",
  "svg:r",
  "svg:spreadMethod",
  "svg:stemh",
  "svg:stemv",
  "svg:stop-color",
  "svg:stop-opacity",
  "svg:strike-through-position",
  "svg:strike-through-thickness",
  "svg:stroke-color",
  "svg:stroke-linecap",
  "svg:stroke-opacity",
  "svg:stroke-width",
  "svg:string",
  "svg:underline-position",
  "svg:underline-thickness",
  "svg:unicode-range",
  "svg:units-per-em",
  "svg:v-alphabetic",
  "svg:v-hanging",
  "svg:v-ideographic",
  "svg:v-mathematical",
  "svg:viewBox",
  "svg:width",
  "svg:widths",
  "svg:x",
  "svg:x-height",
  "svg:x1",
  "svg:x2",
  "svg:y",
  "svg:y1",
  "svg:y2",
  "table:end-cell-address",
  "table:end-x",
  "table:end-y",
  "table:table-background",
  "xlink:actuate",
  "xlink:href",
  "xlink:type",
] as const;

export type DrawingVocabularyElementName = (typeof DRAWING_VOCABULARY_ELEMENTS)[number];
export type DrawingVocabularyAttributeName = (typeof DRAWING_VOCABULARY_ATTRIBUTES)[number];
export type DrawingVocabularyAttributes = Partial<
  Record<DrawingVocabularyAttributeName, string | number | boolean>
>;

export interface OdfDrawingVocabularyElement {
  name: DrawingVocabularyElementName;
  attributes?: DrawingVocabularyAttributes;
  children?: OdfDrawingVocabularyElement[];
  text?: string;
}

export interface OdfMathOptions {
  contentXml: string;
}

export const DRAWING_VOCABULARY_SCHEMA_ELEMENTS = DRAWING_VOCABULARY_ELEMENTS.map((name) => ({
  name,
  attributes: DRAWING_VOCABULARY_ATTRIBUTES,
}));

export const MATH_SCHEMA_ELEMENTS = [{ name: "math:math", attributes: [] }] as const;

const VOCABULARY_ATTRIBUTES = new Set<string>(DRAWING_VOCABULARY_ATTRIBUTES);

function vocabularyError(part: string, path: string, name: string, reason: string): OdfSchemaError {
  return new OdfSchemaError(
    `${part}: ${path}/${name}: ${reason}`,
    part,
    `${path}/${name}`,
    name,
    reason,
  );
}

export function parseDrawingVocabulary(
  root: Element,
  part: string,
  path: string,
): OdfDrawingVocabularyElement | OdfMathOptions {
  if (root.name === "math:math")
    return {
      contentXml: xmlElement(
        root.name,
        root.attributes,
        (root.elements ?? []).map((child) => serializeForeign(child)),
      ),
    };
  if (!DRAWING_VOCABULARY_ELEMENTS.includes(root.name as DrawingVocabularyElementName)) {
    throw vocabularyError(
      part,
      path,
      root.name ?? "unknown",
      "element has no canonical drawing vocabulary mapping",
    );
  }
  return parseVocabularyElement(root, part, path);
}

function parseVocabularyElement(
  element: Element,
  part: string,
  path: string,
): OdfDrawingVocabularyElement {
  const elementPath = `${path}/${element.name}`;
  const attributes: DrawingVocabularyAttributes = {};
  for (const [name, value] of Object.entries(element.attributes ?? {})) {
    if (!VOCABULARY_ATTRIBUTES.has(name))
      throw vocabularyError(
        part,
        path,
        name,
        "attribute has no canonical drawing vocabulary mapping",
      );
    attributes[name as DrawingVocabularyAttributeName] = value;
  }
  return {
    name: element.name as DrawingVocabularyElementName,
    attributes,
    text: textOf(element) || undefined,
    children: (element.elements ?? [])
      .filter((child): child is Element => child.type === "element")
      .map((child) => parseVocabularyElement(child, part, elementPath)),
  };
}

export function drawingVocabularyXml(
  options: OdfDrawingVocabularyElement | OdfMathOptions,
): string {
  if ("contentXml" in options) return options.contentXml;
  return vocabularyElementXml(options);
}

function vocabularyElementXml(element: OdfDrawingVocabularyElement): string {
  return xmlElement(element.name, element.attributes, [
    ...(element.text ? [escapeText(element.text)] : []),
    ...(element.children ?? []).map(vocabularyElementXml),
  ]);
}

function serializeForeign(element: Element): string {
  return xmlElement(element.name ?? "", element.attributes, [
    ...(textOf(element) ? [escapeText(textOf(element))] : []),
    ...(element.elements ?? [])
      .filter((child): child is Element => child.type === "element")
      .map(serializeForeign),
  ]);
}

export function parseMathMarkup(
  contentXml: string,
  part = "content.xml",
  path = "/math:math",
): OdfMathOptions {
  try {
    const root = parse(contentXml, { ignoreDeclaration: true }).elements?.find(
      (element) => element.type === "element",
    );
    if (root?.name !== "math:math") throw new Error("root must be math:math");
    return { contentXml };
  } catch (cause) {
    throw new OdfSchemaError(
      `${part}: ${path}: invalid MathML content`,
      part,
      path,
      "math:math",
      "invalid MathML content",
      { cause },
    );
  }
}
