import type {
  ParagraphDescriptorOptions,
  TextBodyOptions,
  TextRunOptions,
} from "@office-open/core";
import type { PresentationOptions, ShapeOptions, SlideOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { escapeText, metaXml, parseMeta } from "./meta";
import { parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
import { generateOcf, readOcf, readXml } from "./package";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  textOf,
  xmlElement,
} from "./xml";

const MIME = "application/vnd.oasis.opendocument.presentation";
const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
].join(" ");

interface TextProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: "single";
  size?: number;
}

export type OdpOptions = PresentationOptions & { odfExtensions?: OdfXmlNode[] };

export function generateOdp(options: OdpOptions): Uint8Array {
  const styles: string[] = [];
  const size = normalizeSize(options.size);
  const pageLayout = xmlElement("style:page-layout", { "style:name": "PM1" }, [
    xmlElement("style:page-layout-properties", {
      "fo:page-width": emuToLength(size.width),
      "fo:page-height": emuToLength(size.height),
    }),
  ]);
  const pages = (options.slides ?? []).map((slide, index) => slideXml(slide, index + 1, styles));
  const files = {
    "content.xml": contentXml(
      [...pages, ...serializeOdfNodes(options.odfExtensions)].join(""),
      styles,
    ),
    "styles.xml": stylesXml(pageLayout),
    "meta.xml": metaXml(options),
  };
  return generateOcf(MIME, files);
}

export function parseOdp(data: Uint8Array): OdpOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:presentation");
  const stylesDocument = files["styles.xml"] ? readXml(files, "styles.xml") : undefined;
  const pageLayout = childNamed(
    childNamed(childNamed(stylesDocument, "office:automatic-styles"), "style:page-layout"),
    "style:page-layout-properties",
  );
  const width = lengthToEmu(attributeString(pageLayout, "fo:page-width"));
  const height = lengthToEmu(attributeString(pageLayout, "fo:page-height"));
  const rawNodes = parseOdfNodes(body);
  return {
    ...parseMeta(files),
    ...(width && height ? { size: { width, height } } : {}),
    slides: childrenNamed(body, "draw:page").map((page) =>
      parseSlide(page, parseTextStyles(childNamed(content, "office:automatic-styles"))),
    ),
    odfExtensions: rawNodes.filter((node) => node.name !== "draw:page"),
  };
}

function contentXml(pages: string, styles: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
}

function stylesXml(pageLayout: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/><office:automatic-styles>${pageLayout}</office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
}

function normalizeSize(size: PresentationOptions["size"]): { width: number; height: number } {
  if (size === "4:3") return { width: 9144000, height: 6858000 };
  if (typeof size === "object" && size !== null) {
    return {
      width: typeof size.width === "number" ? size.width : (lengthToEmu(size.width) ?? 0),
      height: typeof size.height === "number" ? size.height : (lengthToEmu(size.height) ?? 0),
    };
  }
  return { width: 12192000, height: 6858000 };
}

function slideXml(slide: SlideOptions, index: number, styles: string[]): string {
  const frames = (slide.children ?? []).map((child) =>
    "shape" in child ? shapeXml(child.shape, styles) : "",
  );
  return xmlElement(
    "draw:page",
    { "draw:name": `Slide${index}`, "draw:master-page-name": "Default" },
    frames,
  );
}

function shapeXml(shape: ShapeOptions, styles: string[]): string {
  return xmlElement(
    "draw:frame",
    {
      "draw:name": shape.name,
      "svg:x": emuToLength(shape.x),
      "svg:y": emuToLength(shape.y),
      "svg:width": emuToLength(shape.width),
      "svg:height": emuToLength(shape.height),
    },
    [xmlElement("draw:text-box", undefined, textBodyXml(shape.textBody, styles))],
  );
}

function textBodyXml(body: TextBodyOptions | undefined, styles: string[]): string[] {
  const paragraphs = body?.paragraphs ?? (body?.text !== undefined ? [body.text] : [""]);
  return paragraphs.map((paragraph) => {
    const value = typeof paragraph === "string" ? { text: paragraph } : paragraph;
    const children =
      value.text !== undefined && value.children === undefined
        ? [escapeText(value.text)]
        : (value.children ?? []).map((run) => {
            if (typeof run === "string")
              return xmlElement("text:span", undefined, [escapeText(run)]);
            if (!("text" in run)) return "";
            const properties = textProperties(run);
            const styleName = addTextStyle(properties, styles);
            return xmlElement("text:span", { "text:style-name": styleName }, [
              escapeText(run.text ?? ""),
            ]);
          });
    return xmlElement("text:p", undefined, children);
  });
}

function textProperties(run: TextRunOptions): TextProperties {
  return {
    bold: run.bold,
    italic: run.italic,
    underline: run.underline === "single" ? "single" : undefined,
    size: run.size,
  };
}

function addTextStyle(properties: TextProperties, styles: string[]): string | undefined {
  if (Object.values(properties).every((value) => value === undefined)) return undefined;
  const name = `C${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "text" }, [
      xmlElement("style:text-properties", {
        "fo:font-weight": properties.bold ? "bold" : undefined,
        "fo:font-style": properties.italic ? "italic" : undefined,
        "style:text-underline-style": properties.underline ? "solid" : undefined,
        "fo:font-size": properties.size ? `${properties.size}pt` : undefined,
      }),
    ]),
  );
  return name;
}

function parseSlide(page: Element, textStyles: Map<string, TextProperties>): SlideOptions {
  return {
    children: childrenNamed(page, "draw:frame").map((frame) => ({
      shape: parseShape(frame, textStyles),
    })),
  };
}

function parseShape(frame: Element, textStyles: Map<string, TextProperties>): ShapeOptions {
  const textBox = childNamed(frame, "draw:text-box");
  return {
    name: attributeString(frame, "draw:name"),
    x: lengthToEmu(attributeString(frame, "svg:x")),
    y: lengthToEmu(attributeString(frame, "svg:y")),
    width: lengthToEmu(attributeString(frame, "svg:width")),
    height: lengthToEmu(attributeString(frame, "svg:height")),
    textBody: {
      paragraphs: childrenNamed(textBox, "text:p").map((paragraph) =>
        parseParagraph(paragraph, textStyles),
      ),
    },
  };
}

function parseTextStyles(container: Element | undefined): Map<string, TextProperties> {
  const result = new Map<string, TextProperties>();
  for (const style of childrenNamed(container, "style:style")) {
    const properties = childNamed(style, "style:text-properties");
    const size = attributeString(properties, "fo:font-size");
    result.set(attributeString(style, "style:name") ?? "", {
      bold: attributeString(properties, "fo:font-weight") === "bold",
      italic: attributeString(properties, "fo:font-style") === "italic",
      underline:
        attributeString(properties, "style:text-underline-style") === "solid"
          ? "single"
          : undefined,
      size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
    });
  }
  return result;
}

function parseParagraph(
  paragraph: Element,
  textStyles: Map<string, TextProperties>,
): ParagraphDescriptorOptions {
  const children: (string | TextRunOptions)[] = [];
  for (const child of paragraph.elements ?? []) {
    if (child.type === "text") children.push(String(child.text ?? ""));
    else if (child.name === "text:span") {
      children.push({
        text: textOf(child),
        ...textStyles.get(attributeString(child, "text:style-name") ?? ""),
      });
    }
  }
  if (children.length === 1 && typeof children[0] === "string") return { text: children[0] };
  return children.length ? { children } : { text: "" };
}
