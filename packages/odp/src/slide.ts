import type { ChartSpaceOptions } from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  descendantElements,
  escapeText,
  lengthToEmu,
  parseXmlElement,
  textOf,
  unknownAttributeNames,
  unknownChildElement,
  xmlElement,
  type GraphicStyle,
} from "@office-open/odf";
import type {
  AnimationEntry,
  GroupOptions,
  SlideChild,
  SlideCommentOptions,
  TableOptions,
  TransitionOptions,
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
import { OdpGenerateError, OdpParseError, unknownSlideChild } from "./error";
import type { OdpSlideOptions } from "./options";
import { formsXml, parseForms } from "./presentation";
import {
  assignOdfShapeIds,
  collectOdfShapeIds,
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
  odfShapeIdXml,
  shapeXml,
  toOdfLength,
} from "./shape";
import {
  parseSlideNode,
  slideNodeXml,
  type OdpSlideNodeOptions,
  type OdpSlideNodePayload,
} from "./slide-nodes";
import { addColumnStyle } from "./styles";
import { parseParagraph, textBodyXml, type TextProperties } from "./text-run";

export function slideXml(
  slide: OdpSlideOptions,
  index: number,
  styles: string[],
  images: OdpImage[],
  mediaFiles: OdpImage[],
  embeddedObjects: OdpEmbeddedObjectPool,
  charts: OdpChart[],
): string {
  if (typeof slide.transition === "string") {
    throw new OdpGenerateError(
      "ODP requires a structured slide transition",
      "content.xml",
      "/office:document-content/office:body/office:presentation/draw:page/anim:transitionFilter",
      "transition",
      "raw-transition-not-supported",
    );
  }
  const frames = (slide.children ?? []).map((child) =>
    slideChildXml(child, styles, images, mediaFiles, embeddedObjects, charts),
  );
  for (const node of [...(slide.odpNodes ?? [])].sort(
    (left, right) => left.position - right.position,
  ))
    frames.splice(
      Math.min(node.position, frames.length),
      0,
      slideNodeXml(node.node, styles, charts),
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
      ...(slide.transition === undefined
        ? []
        : Array.isArray(slide.animations) && slide.animations.some(containsTransitionFilter)
          ? []
          : [transitionXml(slide.transition)]),
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

/** ODP extends the canonical slide child union without weakening its PPTX assignment. */
export type OdpSlideChild = SlideChild | OdpSlideNodePayload | { group: OdpGroupOptions };

/** Group options carry typed ODF-only nodes without replacing canonical children. */
export interface OdpGroupOptions extends GroupOptions {
  odpNodes?: OdpSlideNodeOptions[];
}

/** Recursive ODP slide child → ODF dispatcher shared by slides and draw:g groups. */
export function slideChildXml(
  child: OdpSlideChild,
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
    if (child.shape.controlId !== undefined) return controlShapeXml(child.shape, styles);
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
      { "draw:id": odfShapeIdXml(child.group.id), "draw:name": child.group.name },
      (() => {
        const children = child.group.children.map((nested) =>
          slideChildXml(nested, styles, images, mediaFiles, embeddedObjects, charts),
        );
        const odpNodes = "odpNodes" in child.group ? child.group.odpNodes : [];
        for (const node of [...(odpNodes ?? [])].sort(
          (left, right) => left.position - right.position,
        ))
          children.splice(
            Math.min(node.position, children.length),
            0,
            slideNodeXml(node.node, styles, charts),
          );
        return children;
      })(),
    );
  if ("frame" in child || "measure" in child || "pageThumbnail" in child || "scene" in child)
    return slideNodeXml(child, styles, charts);
  if ("rawXml" in child) {
    if (child.rawXml === "") return "";
    const legacyShape = compatibilityLegacyShape(child.rawXml);
    if (legacyShape) return shapeXml(legacyShape, styles);
    const match = /<([A-Za-z_][\w.-]*)(?::([\w.-]+))?/u.exec(child.rawXml);
    const prefix = match?.[2] ? match[2] : undefined;
    if (prefix === undefined || !ODF_RAW_PREFIXES.has(prefix)) {
      throw new OdpGenerateError(
        `Opaque slide node "${prefix}" is not representable as ODF`,
        "content.xml",
        "/office:document-content/office:body/office:presentation/draw:page",
        prefix ?? "rawXml",
        "unsupported-target",
      );
    }
    return child.rawXml;
  }
  return "";
}

const ODF_RAW_PREFIXES = new Set([
  "anim",
  "draw",
  "dr3d",
  "loext",
  "office",
  "presentation",
  "table",
  "text",
]);

type SlideShape = Extract<SlideChild, { shape: unknown }>["shape"];

function localElementName(element: Element): string {
  return element.name?.split(":").pop() ?? "";
}

function compatibilityLegacyShape(rawXml: string): SlideShape | undefined {
  const root = parseXmlElement(rawXml);
  const graphicData = root
    ? descendantElements(root).find((element) => localElementName(element) === "graphicData")
    : undefined;
  if (
    !root ||
    !graphicData ||
    attributeString(graphicData, "uri") !==
      "http://schemas.openxmlformats.org/drawingml/2006/compatibility"
  )
    return undefined;
  const legacyDrawing = descendantElements(root).find(
    (element) => localElementName(element) === "legacyDrawing",
  );
  if (!legacyDrawing) return undefined;
  const nonVisual = childNamed(root, "p:nvGraphicFramePr");
  const drawingProperties = nonVisual ? childNamed(nonVisual, "p:cNvPr") : undefined;
  const transform = childNamed(root, "p:xfrm");
  const offset = transform ? childNamed(transform, "a:off") : undefined;
  const extent = transform ? childNamed(transform, "a:ext") : undefined;
  return {
    ...(attributeNumber(drawingProperties, "id") === undefined
      ? {}
      : { id: attributeNumber(drawingProperties, "id") }),
    ...(attributeString(drawingProperties, "name") === undefined
      ? {}
      : { name: attributeString(drawingProperties, "name") }),
    x: attributeNumber(offset, "x") ?? 0,
    y: attributeNumber(offset, "y") ?? 0,
    width: attributeNumber(extent, "cx") ?? 0,
    height: attributeNumber(extent, "cy") ?? 0,
  };
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
): OdpSlideOptions {
  const normalizedPage = compatibilityPage(page);
  const forms = parseForms(normalizedPage);
  const parsedAnimationRoots = (normalizedPage.elements ?? [])
    .filter((child): child is Element => child.name !== undefined && isAnimationElement(child.name))
    .map(parseAnimation);
  const transition = parseTransition(parsedAnimationRoots);
  const animationRoots = parsedAnimationRoots.filter(
    (entry) => !(transition !== undefined && containsTransitionFilter(entry)),
  );
  const shapeReferences = collectOdfShapeIds(normalizedPage);
  const notes = childNamed(
    childNamed(childNamed(page, "presentation:notes"), "draw:frame"),
    "draw:text-box",
  );
  const notesText = notes ? textOf(childNamed(notes, "text:p")) : undefined;
  const comments = childrenNamed(page, "officeooo:annotation").map(parseAnnotation);
  const children: SlideChild[] = [];
  const odpNodes: OdpSlideNodeOptions[] = [];
  const typedSources = new Set<Element>();
  for (const child of normalizedPage.elements ?? []) {
    if (child.name === "draw:frame") {
      const contentNames = (child.elements ?? [])
        .filter((candidate): candidate is Element => candidate.type === "element")
        .map((candidate) => candidate.name)
        .filter((name): name is string => name !== undefined);
      const frameFallback =
        contentNames.includes("draw:object-ole") ||
        (contentNames.includes("draw:object") &&
          (childNamed(child, "svg:title") !== undefined ||
            childNamed(child, "svg:desc") !== undefined)) ||
        childNamed(child, "office:event-listeners") !== undefined ||
        (contentNames.filter((name) =>
          ["draw:object", "draw:image", "draw:plugin", "draw:text-box", "table:table"].includes(
            name,
          ),
        ).length > 1 &&
          !(
            (contentNames.includes("draw:object") &&
              contentNames.includes("draw:image") &&
              contentNames.length === 2) ||
            (contentNames.includes("draw:plugin") &&
              contentNames.includes("draw:image") &&
              contentNames.length === 2)
          ));
      if (frameFallback) {
        typedSources.add(child);
        const node = parseSlideNode(child);
        const frameChart = parseChartFrame(child, chartPool);
        if (frameChart && "frame" in node) {
          node.frame.contents = [{ kind: "chart", chart: frameChart.chart }];
        }
        odpNodes.push({ position: visibleChildIndex(page, child), node });
        continue;
      }
      const chart = parseChartFrame(child, chartPool);
      if (chart) {
        children.push(chart);
        continue;
      }
      const embeddedObject = parseEmbeddedObjectFrame(
        child,
        embeddedObjects,
        binaries,
        manifestMediaTypes,
        () => nextEmbeddedShapeId.value++,
        textStyles,
      );
      if (embeddedObject) {
        children.push(embeddedObject);
        continue;
      }
      const media = parseMediaFrame(child, binaries, usedBinaryPaths);
      if (media) {
        children.push(media);
        continue;
      }
      const picture = parsePictureFrame(child, binaries, usedBinaryPaths);
      if (picture) {
        children.push(picture);
        continue;
      }
      if (childNamed(child, "draw:text-box")) {
        children.push({ shape: parseShape(child, textStyles) });
        continue;
      }
      typedSources.add(child);
      odpNodes.push({
        position: visibleChildIndex(page, child),
        node: parseSlideNode(child, (frame) => parseChartFrame(frame, chartPool)),
      });
      continue;
    }
    if (child.name === "draw:custom-shape") {
      children.push({ shape: parseCustomSlideShape(child, textStyles, graphicStyles) });
      continue;
    }
    if (child.name === "draw:control") {
      children.push({ shape: parseControlSlideShape(child, textStyles, graphicStyles) });
      continue;
    }
    if (
      child.name === "draw:rect" ||
      child.name === "draw:circle" ||
      child.name === "draw:ellipse" ||
      child.name === "draw:path" ||
      child.name === "draw:polygon" ||
      child.name === "draw:polyline"
    ) {
      children.push({ shape: parsePrimitiveSlideShape(child, textStyles, graphicStyles) });
      continue;
    }
    if (child.name === "draw:line") {
      children.push(parseLine(child));
      continue;
    }
    if (child.name === "draw:g") {
      children.push(
        ...parseGroup(
          child,
          textStyles,
          graphicStyles,
          columnWidths,
          binaries,
          chartPool,
          usedBinaryPaths,
        ),
      );
      continue;
    }
    if (child.name === "draw:connector") {
      children.push(parseConnector(child));
      continue;
    }
    if (child.name === "table:table") {
      children.push(parseSlideTable(child, textStyles, columnWidths));
      continue;
    }
    if (
      child.name === "draw:measure" ||
      child.name === "draw:page-thumbnail" ||
      child.name === "dr3d:scene"
    ) {
      typedSources.add(child);
      odpNodes.push({ position: visibleChildIndex(page, child), node: parseSlideNode(child) });
      continue;
    }
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
  }
  const animations = animationRoots.map((entry) => resolveAnimationTargets(entry, shapeReferences));
  let childIndex = 0;
  for (const source of normalizedPage.elements ?? []) {
    if (source.type !== "element") continue;
    if (
      source.name !== "presentation:notes" &&
      source.name !== "office:forms" &&
      source.name !== "officeooo:annotation" &&
      source.name !== undefined &&
      !isAnimationElement(source.name)
    ) {
      if (typedSources.has(source)) continue;
      assignOdfShapeIds(source, children.slice(childIndex), shapeReferences);
      childIndex++;
    }
  }
  const slideOptions: OdpSlideOptions = {
    ...(attributeString(page, "draw:name") ? { name: attributeString(page, "draw:name") } : {}),
    ...(attributeString(page, "draw:master-page-name")
      ? { master: attributeString(page, "draw:master-page-name") }
      : {}),
    ...(forms !== undefined ? { forms } : {}),
    children,
    ...(odpNodes.length > 0 ? { odpNodes } : {}),
    ...(transition !== undefined ? { transition } : {}),
    ...(animations.length > 0 ? { animations } : {}),
    ...(comments.length > 0 ? { comments } : {}),
    ...(notesText ? { notes: notesText } : {}),
  };
  return slideOptions;
}

function compatibilityPage(page: Element): Element {
  const children = (page.elements ?? []).flatMap((child) => {
    if (child.name !== "mc:AlternateContent") return [child];
    const branches = child.elements ?? [];
    const candidates = [
      ...branches.filter((branch) => branch.name === "mc:Choice"),
      ...branches.filter((branch) => branch.name === "mc:Fallback"),
    ];
    const known = candidates.flatMap((branch) =>
      (branch.elements ?? []).filter((content) => isKnownCompatibilityChild(content.name)),
    );
    const selected = (known[0] ?? candidates[0]?.elements?.[0]) as Element | undefined;
    if (selected === undefined) {
      throw unknownSlideChild(
        child,
        "/office:document-content/office:body/office:presentation/draw:page/mc:AlternateContent",
      );
    }
    return [selected];
  });
  return { ...page, elements: children };
}

function visibleChildIndex(page: Element, child: Element): number {
  let index = 0;
  for (const candidate of page.elements ?? []) {
    if (candidate.type !== "element") continue;
    if (
      candidate.name === "presentation:notes" ||
      candidate.name === "office:forms" ||
      candidate.name === "officeooo:annotation" ||
      (candidate.name !== undefined && isAnimationElement(candidate.name))
    )
      continue;
    if (candidate === child) return index;
    index++;
  }
  return index;
}

function isKnownCompatibilityChild(name: string | undefined): boolean {
  if (name === undefined) return false;
  return (
    name.startsWith("draw:") ||
    name.startsWith("anim:") ||
    name === "table:table" ||
    name === "presentation:notes" ||
    name === "office:forms" ||
    name === "officeooo:annotation"
  );
}

function findTransitionFilter(entry: AnimationEntry): AnimationEntry | undefined {
  if (entry.behavior === "transitionFilter") return entry;
  for (const child of entry.children ?? []) {
    const transition = findTransitionFilter(child);
    if (transition !== undefined) return transition;
  }
  return undefined;
}

const TRANSITION_TYPES = {
  barWipe: "wipe",
  blindsWipe: "blinds",
  checkerBoardWipe: "checker",
  dissolve: "dissolve",
  ellipseWipe: "circle",
  fade: "fade",
  fanWipe: "wedge",
  fourBoxWipe: "plus",
  irisWipe: "diamond",
  pushWipe: "push",
  randomBarWipe: "randomBar",
  slideWipe: "cover",
} as const satisfies Record<string, TransitionOptions["type"]>;

const SMIL_TRANSITION_TYPES: Partial<Record<NonNullable<TransitionOptions["type"]>, string>> = {
  blinds: "blindsWipe",
  checker: "checkerBoardWipe",
  circle: "ellipseWipe",
  cover: "slideWipe",
  diamond: "irisWipe",
  dissolve: "dissolve",
  fade: "fade",
  plus: "fourBoxWipe",
  push: "pushWipe",
  random: "random",
  randomBar: "randomBarWipe",
  wedge: "fanWipe",
  wipe: "barWipe",
  zoom: "zoom",
};

function containsTransitionFilter(entry: AnimationEntry): boolean {
  if (entry.behavior === "transitionFilter") return true;
  return (entry.children ?? []).some(containsTransitionFilter);
}

function parseTransition(entries: readonly AnimationEntry[]): TransitionOptions | undefined {
  const filter = entries.map(findTransitionFilter).find(Boolean);
  if (filter === undefined) return undefined;
  const type = TRANSITION_TYPES[(filter.transitionType as keyof typeof TRANSITION_TYPES) ?? ""];
  if (type === undefined) {
    throw new OdpParseError(
      `unsupported slide transition type: ${filter.transitionType}`,
      "content.xml",
      "/office:document-content/office:body/office:presentation/draw:page/anim:transitionFilter",
      "anim:transitionFilter",
      "unsupported-smil-transition-type",
    );
  }
  const subtype = filter.transitionSubType;
  return {
    type,
    ...(filter.duration !== undefined && typeof filter.duration === "number"
      ? { duration: filter.duration }
      : filter.duration === "indefinite"
        ? {}
        : {}),
    ...(subtype?.includes("horizontal") || subtype === "across" ? { orient: "horizontal" } : {}),
    ...(subtype?.includes("vertical") || subtype === "down" ? { orient: "vertical" } : {}),
    ...(subtype === "fadeOverColor" || subtype === "rectangle" ? { thruBlk: true } : {}),
    ...(subtype === "fromRight" ? { direction: "left" } : {}),
    ...(subtype === "fromLeft" ? { direction: "right" } : {}),
    ...(subtype === "fromBottom" ? { direction: "up" } : {}),
    ...(subtype === "fromTop" ? { direction: "down" } : {}),
  } satisfies TransitionOptions;
}

function transitionXml(options: TransitionOptions): string {
  const smilType = options.type === undefined ? undefined : SMIL_TRANSITION_TYPES[options.type];
  if (!smilType) {
    throw new OdpGenerateError(
      `ODP cannot represent slide transition type: ${options.type}`,
      "content.xml",
      "/office:document-content/office:body/office:presentation/draw:page/anim:transitionFilter",
      options.type ?? "(missing)",
      "unsupported-transition-type",
    );
  }
  const subtype =
    options.type === "fade"
      ? options.thruBlk
        ? "fadeOverColor"
        : "crossfade"
      : options.orient !== undefined
        ? options.orient === "horizontal"
          ? "horizontal"
          : "vertical"
        : options.direction === "left"
          ? "fromRight"
          : options.direction === "right"
            ? "fromLeft"
            : options.direction === "up"
              ? "fromBottom"
              : options.direction === "down"
                ? "fromTop"
                : undefined;
  const filter = xmlElement("anim:transitionFilter", {
    "smil:type": smilType,
    ...(subtype ? { "smil:subtype": subtype } : {}),
    ...(options.duration !== undefined ? { "smil:dur": `${options.duration}ms` } : {}),
  });
  return xmlElement("anim:par", { "presentation:node-type": "timing-root" }, [filter]);
}

function resolveAnimationTargets(
  entry: AnimationEntry,
  references: Map<string, number>,
): AnimationEntry {
  const rawId = entry.shapeName === undefined ? undefined : references.get(entry.shapeName);
  const resolved: AnimationEntry =
    rawId === undefined ? entry : { ...entry, shapeId: rawId, shapeName: undefined };
  return resolved.children?.length
    ? {
        ...resolved,
        children: resolved.children.map((child) => resolveAnimationTargets(child, references)),
      }
    : resolved;
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
    if (width) result.set(attributeString(style, "style:name") ?? "", lengthToEmu(width) ?? 0);
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
  return xmlElement("table:table", { "table:name": table.name }, [...columns, ...rows]);
}

/** Parses a slide table:table back to a pptx table child. */
export function parseSlideTable(
  element: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
): { table: TableOptions } {
  const rows = childrenNamed(element, "table:table-row").map((row) => ({
    cells: parseTableRowCells(row, textStyles),
  }));
  return {
    table: {
      ...(attributeString(element, "table:name")
        ? { name: attributeString(element, "table:name") }
        : {}),
      columnWidths: childrenNamed(element, "table:table-column").map(
        (column) => columnWidths.get(attributeString(column, "table:style-name") ?? "") ?? 0,
      ),
      rows,
    },
  };
}

function parseTableRowCells(
  row: Element,
  textStyles: Map<string, TextProperties>,
): TableOptions["rows"][number]["cells"] {
  const children = (row.elements ?? []).filter(
    (child): child is Element => child.type === "element",
  );
  const cells: TableOptions["rows"][number]["cells"] = [];
  for (let index = 0; index < children.length; index += 1) {
    const cell = children[index]!;
    if (cell.name !== "table:table-cell") throw unknownSlideChild(cell, "/draw:page/table:table");
    const span = attributeNumber(cell, "table:number-columns-spanned") ?? 1;
    cells.push({
      columnSpan: attributeNumber(cell, "table:number-columns-spanned"),
      children: childrenNamed(cell, "text:p").map((paragraph) =>
        parseParagraph(paragraph, textStyles),
      ),
    });
    const coveredCount = span - 1;
    for (let offset = 1; offset <= coveredCount; offset += 1) {
      if (children[index + offset]?.name !== "table:covered-table-cell") {
        throw new OdpParseError(
          "Invalid ODP table covered-cell sequence",
          "content.xml",
          "/draw:page/table:table/table:table-row",
          cell.name ?? "",
          "covered-cell-count-does-not-match-column-span",
        );
      }
    }
    if (children[index + coveredCount + 1]?.name === "table:covered-table-cell") {
      throw new OdpParseError(
        "Invalid ODP table covered-cell sequence",
        "content.xml",
        "/draw:page/table:table/table:table-row",
        cell.name ?? "",
        "covered-cell-count-does-not-match-column-span",
      );
    }
    index += coveredCount;
  }
  return cells;
}
