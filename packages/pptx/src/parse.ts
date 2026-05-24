import type { ParsedDocument } from "@office-open/core";
import { parseDocument } from "@office-open/core";
import type { Element } from "@office-open/xml";
import { attr } from "@office-open/xml";

import { ParseContext } from "./parse/context";
import { parseSlide } from "./parse/slide";

export { parseDocument };

import type { ISlideOptions } from "./file/file";

/**
 * All part paths extracted from the PPTX package.
 * Field names correspond directly to the OOXML directory structure.
 */
export interface PptxPartRefs {
  /** ppt/theme/themeN.xml */
  themes: string[];
  /** ppt/notesMasters/notesMasterN.xml */
  notesMasters: string[];
  /** ppt/commentAuthors.xml */
  commentAuthors?: string;
  /** ppt/comments/commentN.xml (from slide rels) */
  comments: string[];
  /** ppt/charts/chartN.xml (from slide rels) */
  charts: string[];
  /** ppt/diagrams/dataN.xml (from slide rels) */
  diagramData: string[];
  /** ppt/media/* (all media files) */
  media: string[];
}

export interface PptxDocument {
  doc: ParsedDocument;
  /** ppt/presentation.xml root element (p:presentation) */
  presentation?: Element;
  /** ppt/slides/slideN.xml */
  slides: string[];
  /** ppt/slideMasters/slideMasterN.xml */
  slideMasters: string[];
  /** ppt/slideLayouts/slideLayoutN.xml */
  slideLayouts: string[];
  /** ppt/notesSlides/notesSlideN.xml */
  notesSlides: string[];
  partRefs: PptxPartRefs;
  /** ppt/presProps.xml */
  presProps?: string;
  /** ppt/viewProps.xml */
  viewProps?: string;
  /** ppt/tableStyles.xml */
  tableStyles?: string;
  /** docProps/core.xml */
  coreProps?: string;
  /** docProps/app.xml */
  appProps?: string;
}

function resolveRelsPath(target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  if (target.startsWith("../")) return target.replace("../", "ppt/");
  return `ppt/${target}`;
}

function sortByNumber(paths: string[]): string[] {
  return paths.sort((a, b) => {
    const numA = parseInt(a.match(/(\d+)/)?.[1] ?? "0", 10);
    const numB = parseInt(b.match(/(\d+)/)?.[1] ?? "0", 10);
    return numA - numB;
  });
}

function xmlKeys(keys: string[]): string[] {
  return keys.filter((k) => k.endsWith(".xml"));
}

function parseRootRels(doc: ParsedDocument): { coreProps?: string; appProps?: string } {
  const relsEl = doc.get("_rels/.rels");
  if (!relsEl) return {};

  let coreProps: string | undefined;
  let appProps: string | undefined;

  for (const child of relsEl.elements ?? []) {
    if (child.name !== "Relationship") continue;
    const type = attr(child, "Type") ?? "";
    const target = attr(child, "Target") ?? "";
    if (!target) continue;

    const path = target.startsWith("/") ? target.slice(1) : target;

    if (type.includes("/core-properties")) {
      coreProps = path;
    } else if (type.includes("/extended-properties")) {
      appProps = path;
    }
  }

  return { coreProps, appProps };
}

function parseSlideRels(doc: ParsedDocument, slidePaths: string[], refs: PptxPartRefs): void {
  for (const slidePath of slidePaths) {
    const parts = slidePath.split("/");
    const fileName = parts.pop()!;
    const relsPath = `${parts.join("/")}/_rels/${fileName}.rels`;

    const relsEl = doc.get(relsPath);
    if (!relsEl) continue;

    for (const child of relsEl.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (!target) continue;

      const path = resolveRelsPath(target);

      if (type.includes("/comments") && !type.includes("commentAuthors")) {
        if (!refs.comments.includes(path)) refs.comments.push(path);
      } else if (type.includes("/chart")) {
        if (!refs.charts.includes(path)) refs.charts.push(path);
      } else if (type.includes("/diagramData")) {
        if (!refs.diagramData.includes(path)) refs.diagramData.push(path);
      } else if (type.includes("/image") || type.includes("/video") || type.includes("/media")) {
        if (!refs.media.includes(path)) refs.media.push(path);
      }
    }
  }
}

export function parsePptx(data: Uint8Array): PptxDocument {
  const doc = parseDocument(data);

  const presentation = doc.get("ppt/presentation.xml");

  const relsXml = doc.get("ppt/_rels/presentation.xml.rels");
  const slides: string[] = [];
  const slideMasters: string[] = [];
  const themes: string[] = [];
  const notesMasters: string[] = [];
  let presProps: string | undefined;
  let viewProps: string | undefined;
  let tableStyles: string | undefined;
  let commentAuthors: string | undefined;

  if (relsXml) {
    for (const child of relsXml.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (!target) continue;

      const path = resolveRelsPath(target);

      if (type.includes("/slideMaster")) {
        slideMasters.push(path);
      } else if (
        type.includes("/slide") &&
        !type.includes("slideLayout") &&
        !type.includes("slideMaster")
      ) {
        slides.push(path);
      } else if (type.includes("/theme")) {
        themes.push(path);
      } else if (type.includes("/notesMaster")) {
        notesMasters.push(path);
      } else if (type.includes("/presProps")) {
        presProps = path;
      } else if (type.includes("/viewProps")) {
        viewProps = path;
      } else if (type.includes("/tableStyles")) {
        tableStyles = path;
      } else if (type.includes("/commentAuthors")) {
        commentAuthors = path;
      }
    }
  }

  sortByNumber(slides);
  sortByNumber(slideMasters);
  sortByNumber(themes);
  sortByNumber(notesMasters);

  const slideLayouts = sortByNumber(xmlKeys(doc.keys("ppt/slideLayouts/")));
  const notesSlides = sortByNumber(xmlKeys(doc.keys("ppt/notesSlides/")));

  const partRefs: PptxPartRefs = {
    themes,
    notesMasters,
    commentAuthors,
    comments: [],
    charts: [],
    diagramData: [],
    media: doc.keys("ppt/media/"),
  };

  parseSlideRels(doc, slides, partRefs);
  sortByNumber(partRefs.comments);
  sortByNumber(partRefs.charts);
  sortByNumber(partRefs.diagramData);

  const { coreProps, appProps } = parseRootRels(doc);

  return {
    doc,
    presentation,
    slides,
    slideMasters,
    slideLayouts,
    notesSlides,
    partRefs,
    presProps,
    viewProps,
    tableStyles,
    coreProps,
    appProps,
  };
}

/**
 * Parse a single slide's relationship file into a Map<rId, path>.
 */
function parseSlideRelMap(doc: ParsedDocument, slidePath: string): Map<string, string> {
  const rels = new Map<string, string>();
  const parts = slidePath.split("/");
  const fileName = parts.pop()!;
  const relsPath = `${parts.join("/")}/_rels/${fileName}.rels`;

  const relsEl = doc.get(relsPath);
  if (!relsEl) return rels;

  for (const child of relsEl.elements ?? []) {
    if (child.name !== "Relationship") continue;
    const id = attr(child, "Id") ?? "";
    const target = attr(child, "Target") ?? "";
    if (!id || !target) continue;
    rels.set(id, resolveRelsPath(target));
  }

  return rels;
}

/**
 * Read a .pptx file and convert it into ISlideOptions[].
 *
 * This is the main public API for parsing PPTX files.
 * The returned options can be passed directly to `new Presentation({ slides })`
 * to recreate the presentation.
 *
 * @param data - Raw bytes of a .pptx file
 * @returns Array of slide options
 */
export function readPresentation(data: Uint8Array): ISlideOptions[] {
  const pptx = parsePptx(data);
  const result: ISlideOptions[] = [];

  for (const slidePath of pptx.slides) {
    const slideEl = pptx.doc.get(slidePath);
    if (!slideEl) continue;

    const slideRels = parseSlideRelMap(pptx.doc, slidePath);
    const ctx = new ParseContext(pptx, slideRels);
    result.push(parseSlide(slideEl, ctx));
  }

  return result;
}
