import type { ChartSpaceOptions } from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  escapeText,
  lengthToEmu,
  textOf,
  xmlElement,
  type GraphicStyle,
} from "@office-open/odf";
import type { SlideChild, SlideOptions, TableOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import {
  chartFrameXml,
  parseChartFrame,
  parsePictureFrame,
  pictureFrameXml,
  type OdpChart,
  type OdpImage,
} from "./drawing";
import { formsXml, parseForms, unknownSlideChild } from "./presentation";
import {
  connectorXml,
  customShapeXml,
  lineXml,
  parseConnector,
  parseCustomSlideShape,
  parseGroup,
  parseLine,
  parseShape,
  shapeXml,
} from "./shape";
import { addColumnStyle } from "./styles";
import { parseParagraph, textBodyXml, type TextProperties } from "./text-run";

export function slideXml(
  slide: SlideOptions,
  index: number,
  styles: string[],
  images: OdpImage[],
  charts: OdpChart[],
): string {
  const frames = (slide.children ?? []).map((child) =>
    slideChildXml(child, styles, images, charts),
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
    { "draw:name": `Slide${index}`, "draw:master-page-name": "Default" },
    [
      ...(slide.forms?.length ? [formsXml(slide.forms)] : []),
      ...frames,
      ...(notesXml ? [notesXml] : []),
    ],
  );
}

/** Recursive SlideChild → ODF dispatcher shared by slides and draw:g groups. */
export function slideChildXml(
  child: SlideChild,
  styles: string[],
  images: OdpImage[],
  charts: OdpChart[],
): string {
  if ("shape" in child) {
    if (child.shape.properties?.geometry !== undefined) return customShapeXml(child.shape, styles);
    return shapeXml(child.shape, styles);
  }
  if ("table" in child) return slideTableXml(child.table, styles);
  if ("connector" in child) return connectorXml(child.connector);
  if ("line" in child) return lineXml(child.line);
  if ("picture" in child) return pictureFrameXml(child.picture, images);
  if ("chart" in child) return chartFrameXml(child.chart, charts);
  if ("group" in child)
    return xmlElement(
      "draw:g",
      { "draw:name": child.group.name },
      child.group.children.map((nested) => slideChildXml(nested, styles, images, charts)),
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
): SlideOptions {
  const forms = parseForms(page);
  const notes = childNamed(
    childNamed(childNamed(page, "presentation:notes"), "draw:frame"),
    "draw:text-box",
  );
  const notesText = notes ? textOf(childNamed(notes, "text:p")) : undefined;
  const slideOptions = {
    ...(forms.length > 0 ? { forms } : {}),
    children:
      page.elements?.flatMap((child): SlideChild[] => {
        if (child.name === "draw:frame") {
          const chart = parseChartFrame(child, chartPool);
          if (chart) return [chart];
          const picture = parsePictureFrame(child, binaries);
          if (picture) return [picture];
          if (childNamed(child, "draw:text-box")) return [{ shape: parseShape(child, textStyles) }];
          throw unknownSlideChild(
            child,
            "/office:document-content/office:body/office:presentation/draw:page",
          );
        }
        if (child.name === "draw:custom-shape")
          return [{ shape: parseCustomSlideShape(child, textStyles, graphicStyles) }];
        if (child.name === "draw:line") return [parseLine(child)];
        if (child.name === "draw:g")
          return parseGroup(child, textStyles, columnWidths, binaries, chartPool);
        if (child.name === "draw:connector") return [parseConnector(child)];
        if (child.name === "table:table") return [parseSlideTable(child, textStyles, columnWidths)];
        if (child.name !== "presentation:notes" && child.name !== "office:forms")
          throw unknownSlideChild(
            child,
            "/office:document-content/office:body/office:presentation/draw:page",
          );
        return [];
      }) ?? [],
    ...(notesText ? { notes: notesText } : {}),
  };
  return slideOptions;
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
