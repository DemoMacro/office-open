import type { ChartSpaceOptions } from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  descendantElements,
  escapeText,
  lengthToEmu,
  textOf,
  unknownAttributeNames,
  unknownChildElement,
  xmlElement,
  type GraphicStyle,
} from "@office-open/odf";
import type {
  SlideChild,
  SlideCommentOptions,
  SlideOptions,
  TableOptions,
} from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { animationsXml, isAnimationElement, parseAnimation } from "./animation";
import {
  chartFrameXml,
  parseChartFrame,
  mediaFrameXml,
  parseMediaFrame,
  parsePictureFrame,
  pictureFrameXml,
  type OdpChart,
  type OdpImage,
} from "./drawing";
import {
  embeddedObjectXml,
  parseEmbeddedObjectFrame,
  type OdpEmbeddedObjectPool,
} from "./embedded-frames";
import { formsXml, parseForms, unknownSlideChild } from "./presentation";
import {
  controlShapeXml,
  connectorXml,
  customShapeXml,
  lineXml,
  parseConnector,
  parseControlSlideShape,
  parseCustomSlideShape,
  parseGroup,
  parseLine,
  parsePrimitiveSlideShape,
  parseShape,
  shapeXml,
  toOdfLength,
} from "./shape";
import { addColumnStyle } from "./styles";
import { parseParagraph, textBodyXml, type TextProperties } from "./text-run";

export function slideXml(
  slide: SlideOptions,
  index: number,
  styles: string[],
  images: OdpImage[],
  mediaFiles: OdpImage[],
  embeddedObjects: OdpEmbeddedObjectPool,
  charts: OdpChart[],
): string {
  const frames = (slide.children ?? []).map((child) =>
    slideChildXml(child, styles, images, mediaFiles, embeddedObjects, charts),
  );
  const notes = typeof slide.notes === "string" ? slide.notes : slide.notes?.text;
  const notesXml = notes
    ? xmlElement("presentation:notes", undefined, [
        xmlElement("draw:frame", undefined, [
          xmlElement("draw:text-box", undefined, [
            xmlElement("text:p", undefined, [escapeText(notes)]),
          ]),
        ]),
      ])
    : undefined;
  return xmlElement(
    "draw:page",
    {
      "draw:name": slide.name ?? `Slide${index}`,
      "draw:master-page-name": slide.master ?? "Default",
    },
    [
      ...(slide.forms === undefined ? [] : [formsXml(slide.forms)]),
      ...(Array.isArray(slide.animations) && animationsXml(slide.animations)
        ? [animationsXml(slide.animations)!]
        : []),
      ...frames,
      ...(slide.comments ?? []).map(annotationXml),
      ...(notesXml ? [notesXml] : []),
    ],
  );
}

function annotationXml(comment: SlideCommentOptions): string {
  return xmlElement(
    "officeooo:annotation",
    {
      "svg:x": toOdfLength(comment.x),
      "svg:y": toOdfLength(comment.y),
      ...(comment.width === undefined ? {} : { "svg:width": toOdfLength(comment.width) }),
      ...(comment.height === undefined ? {} : { "svg:height": toOdfLength(comment.height) }),
    },
    [
      xmlElement("dc:creator", undefined, [escapeText(comment.author)]),
      ...(comment.initials
        ? [xmlElement("loext:sender-initials", undefined, [escapeText(comment.initials)])]
        : []),
      ...(comment.date ? [xmlElement("dc:date", undefined, [escapeText(comment.date)])] : []),
      ...comment.text
        .split("\n")
        .map((text) => xmlElement("text:p", undefined, [escapeText(text)])),
    ],
  );
}

/** Recursive SlideChild → ODF dispatcher shared by slides and draw:g groups. */
export function slideChildXml(
  child: SlideChild,
  styles: string[],
  images: OdpImage[],
  mediaFiles: OdpImage[],
  embeddedObjects: OdpEmbeddedObjectPool,
  charts: OdpChart[],
): string {
  if ("shape" in child) {
    const object = embeddedObjects
      .values()
      .find((candidate) => candidate.frames?.some((frame) => frame.shapeId === child.shape.id));
    const frame = object?.frames?.find((candidate) => candidate.shapeId === child.shape.id);
    if (object && frame) return embeddedObjectXml(child.shape, object, frame);
    if (child.shape.odfControlId !== undefined) return controlShapeXml(child.shape, styles);
    if (
      child.shape.properties?.geometry !== undefined ||
      child.shape.properties?.customGeometry !== undefined
    )
      return customShapeXml(child.shape, styles);
    return shapeXml(child.shape, styles);
  }
  if ("table" in child) return slideTableXml(child.table, styles);
  if ("connector" in child) return connectorXml(child.connector);
  if ("line" in child) return lineXml(child.line);
  if ("picture" in child) return pictureFrameXml(child.picture, images);
  if ("video" in child) return mediaFrameXml(child.video, mediaFiles, images);
  if ("audio" in child) return mediaFrameXml(child.audio, mediaFiles, images);
  if ("chart" in child) return chartFrameXml(child.chart, charts);
  if ("group" in child)
    return xmlElement(
      "draw:g",
      { "draw:name": child.group.name },
      child.group.children.map((nested) =>
        slideChildXml(nested, styles, images, mediaFiles, embeddedObjects, charts),
      ),
    );
  return "";
}

export function parseSlide(
  page: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
  binaries: Record<string, Uint8Array>,
  graphicStyles: Map<string, GraphicStyle>,
  chartPool: Map<string, ChartSpaceOptions>,
  usedBinaryPaths: Set<string>,
  embeddedObjects: OdpEmbeddedObjectPool,
  manifestMediaTypes: Map<string, string | undefined>,
  nextEmbeddedShapeId: { value: number },
): SlideOptions {
  const forms = parseForms(page);
  const animations = (page.elements ?? [])
    .filter((child): child is Element => child.name !== undefined && isAnimationElement(child.name))
    .map(parseAnimation);
  const notes = childNamed(
    childNamed(childNamed(page, "presentation:notes"), "draw:frame"),
    "draw:text-box",
  );
  const notesText = notes ? textOf(childNamed(notes, "text:p")) : undefined;
  const comments = childrenNamed(page, "officeooo:annotation").map(parseAnnotation);
  const slideOptions = {
    ...(attributeString(page, "draw:name") ? { name: attributeString(page, "draw:name") } : {}),
    ...(attributeString(page, "draw:master-page-name")
      ? { master: attributeString(page, "draw:master-page-name") }
      : {}),
    ...(forms !== undefined ? { forms } : {}),
    children:
      page.elements?.flatMap((child): SlideChild[] => {
        if (child.name === "draw:frame") {
          const chart = parseChartFrame(child, chartPool);
          if (chart) return [chart];
          const embeddedObject = parseEmbeddedObjectFrame(
            child,
            embeddedObjects,
            binaries,
            manifestMediaTypes,
            () => nextEmbeddedShapeId.value++,
            textStyles,
          );
          if (embeddedObject) return [embeddedObject];
          const media = parseMediaFrame(child, binaries, usedBinaryPaths);
          if (media) return [media];
          if (childNamed(child, "draw:object-ole")) return [];
          const picture = parsePictureFrame(child, binaries, usedBinaryPaths);
          if (picture) return [picture];
          if (childNamed(child, "draw:text-box")) return [{ shape: parseShape(child, textStyles) }];
          throw unknownSlideChild(
            child,
            "/office:document-content/office:body/office:presentation/draw:page",
          );
        }
        if (child.name === "draw:custom-shape")
          return [{ shape: parseCustomSlideShape(child, textStyles, graphicStyles) }];
        if (child.name === "draw:control")
          return [{ shape: parseControlSlideShape(child, textStyles, graphicStyles) }];
        if (
          child.name === "draw:rect" ||
          child.name === "draw:circle" ||
          child.name === "draw:ellipse" ||
          child.name === "draw:path" ||
          child.name === "draw:polygon" ||
          child.name === "draw:polyline"
        )
          return [{ shape: parsePrimitiveSlideShape(child, textStyles, graphicStyles) }];
        if (child.name === "draw:line") return [parseLine(child)];
        if (child.name === "draw:g")
          return parseGroup(
            child,
            textStyles,
            graphicStyles,
            columnWidths,
            binaries,
            chartPool,
            usedBinaryPaths,
          );
        if (child.name === "draw:connector") return [parseConnector(child)];
        if (child.name === "table:table") return [parseSlideTable(child, textStyles, columnWidths)];
        if (
          child.name !== "presentation:notes" &&
          child.name !== "office:forms" &&
          child.name !== "officeooo:annotation" &&
          !(child.name !== undefined && isAnimationElement(child.name))
        )
          throw unknownSlideChild(
            child,
            "/office:document-content/office:body/office:presentation/draw:page",
          );
        return [];
      }) ?? [],
    ...(animations.length > 0 ? { animations } : {}),
    ...(comments.length > 0 ? { comments } : {}),
    ...(notesText ? { notes: notesText } : {}),
  };
  return slideOptions;
}

function parseAnnotation(element: Element): SlideCommentOptions {
  const unknownAttribute = unknownAttributeNames(element, [
    "svg:x",
    "svg:y",
    "svg:width",
    "svg:height",
    "dc:creator",
    "meta:creator-initials",
    "dc:date",
    "loext:sender-initials",
  ])[0];
  if (unknownAttribute) throw unknownSlideChild(element, "/draw:page/officeooo:annotation");
  const unknownChild = unknownChildElement(element, [
    "dc:creator",
    "meta:creator-initials",
    "dc:date",
    "loext:sender-initials",
    "text:p",
  ]);
  if (unknownChild) throw unknownSlideChild(element, "/draw:page/officeooo:annotation");
  const author = textOf(childNamed(element, "dc:creator"));
  if (!author) throw unknownSlideChild(element, "/draw:page/officeooo:annotation");
  const childByLocalName = (name: string): Element | undefined =>
    childNamed(element, name) ??
    descendantElements(element).find((child) => child.name?.split(":").pop() === name);
  const width = attributeString(element, "svg:width");
  const height = attributeString(element, "svg:height");
  return {
    author,
    text: childrenNamed(element, "text:p").map(textOf).join("\n"),
    x: lengthToEmu(attributeString(element, "svg:x")) ?? 0,
    y: lengthToEmu(attributeString(element, "svg:y")) ?? 0,
    ...(textOf(childByLocalName("date")) ? { date: textOf(childByLocalName("date")) } : {}),
    ...(textOf(childByLocalName("sender-initials"))
      ? { initials: textOf(childByLocalName("sender-initials")) }
      : {}),
    ...(width ? { width: lengthToEmu(width) ?? 0 } : {}),
    ...(height ? { height: lengthToEmu(height) ?? 0 } : {}),
  };
}

/** Maps table-column style names to twip widths for slide tables. */
export function parseColumnWidths(container: Element | undefined): Map<string, number> {
  const result = new Map<string, number>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "table-column") continue;
    const width = attributeString(
      childNamed(style, "style:table-column-properties"),
      "style:column-width",
    );
    if (width?.endsWith("cm")) {
      result.set(attributeString(style, "style:name") ?? "", lengthToEmu(width) ?? 0);
    }
  }
  return result;
}

/** Serializes a pptx table as an ODF table:table (cells render text paragraphs). */
export function slideTableXml(table: TableOptions, styles: string[]): string {
  const columns = (table.columnWidths ?? []).map((width) =>
    xmlElement("table:table-column", { "table:style-name": addColumnStyle(width, styles) }),
  );
  const rows = table.rows.map((row) =>
    xmlElement(
      "table:table-row",
      undefined,
      row.cells.map((cell) => {
        const content = textBodyXml(
          cell.children ? { paragraphs: cell.children } : { text: cell.text ?? "" },
          styles,
        );
        const span = cell.columnSpan;
        const xml = xmlElement(
          "table:table-cell",
          { "office:value-type": "string", "table:number-columns-spanned": span },
          content,
        );
        const covered = Array.from({ length: Math.max(0, (span ?? 1) - 1) }, () =>
          xmlElement("table:covered-table-cell"),
        );
        return xml + covered.join("");
      }),
    ),
  );
  return xmlElement("table:table", { "table:name": `Table${styles.length + 1}` }, [
    ...columns,
    ...rows,
  ]);
}

/** Parses a slide table:table back to a pptx table child. */
export function parseSlideTable(
  element: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
): { table: TableOptions } {
  return {
    table: {
      columnWidths: childrenNamed(element, "table:table-column").map(
        (column) => columnWidths.get(attributeString(column, "table:style-name") ?? "") ?? 0,
      ),
      rows: childrenNamed(element, "table:table-row").map((row) => ({
        cells: childrenNamed(row, "table:table-cell").map((cell) => ({
          columnSpan: attributeNumber(cell, "table:number-columns-spanned"),
          children: childrenNamed(cell, "text:p").map((paragraph) =>
            parseParagraph(paragraph, textStyles),
          ),
        })),
      })),
    },
  };
}
