import type { ParsedDocument } from "@office-open/core";
import { parseArchive, parseCorePropsElement, convertEmuToPixels } from "@office-open/core";
import type { Element } from "@office-open/xml";
import { attr, attrNum, findChild, textOf } from "@office-open/xml";
import type { DataType } from "undio";
import { toUint8Array } from "undio";

import { ParseContext } from "./parse/context";
import { parseSlide } from "./parse/slide";
import { parseBackground } from "./parse/slide";
import { parseSlideLayoutType } from "./parse/slide-layout";
import { parseTheme } from "./parse/theme";

export { parseArchive };

import type { MasterDefinition, SlideOptions, PresentationOptions } from "./file/file";
import type { SlideLayoutType } from "./file/slide-layout/slide-layout";

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

export function parsePptx(data: DataType): PptxDocument {
  const uint8 = toUint8Array(data);
  const doc = parseArchive(uint8);

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
 * Build a map from each path to the rel target matching a predicate.
 */
function resolveRelTargets(
  doc: ParsedDocument,
  paths: readonly string[],
  predicate: (target: string) => boolean,
): Map<string, string> {
  const map = new Map<string, string>();
  for (const path of paths) {
    for (const target of parseSlideRelMap(doc, path).values()) {
      if (predicate(target)) map.set(path, target);
    }
  }
  return map;
}

/**
 * Parse a .pptx file and convert it into PresentationOptions.
 *
 * This is the main public API for parsing PPTX files.
 * The returned options can be passed directly to `new Presentation(parsed)`
 * to recreate the presentation.
 *
 * @param data - Raw bytes of a .pptx file
 * @returns Parsed presentation options
 */
export function parsePresentation(data: DataType): PresentationOptions {
  const pptx = parsePptx(data);
  const opts: Record<string, unknown> = {};

  // 1. Parse slide size from p:sldSz
  if (pptx.presentation) {
    const sldSz = findChild(pptx.presentation, "p:sldSz");
    if (sldSz) {
      const cx = attrNum(sldSz, "cx");
      const cy = attrNum(sldSz, "cy");
      if (cx === 12192000 && cy === 6858000) {
        opts.size = "16:9";
      } else if (cx === 9144000 && cy === 6858000) {
        opts.size = "4:3";
      } else if (cx && cy) {
        opts.size = { width: convertEmuToPixels(cx), height: convertEmuToPixels(cy) };
      }
    }
  }

  // 2. Parse core properties
  if (pptx.coreProps) {
    const corePropsEl = pptx.doc.get(pptx.coreProps);
    if (corePropsEl) {
      const cp = parseCorePropsElement(corePropsEl);
      if (cp.title) opts.title = cp.title;
      if (cp.subject) opts.subject = cp.subject;
      if (cp.creator) opts.creator = cp.creator;
      if (cp.keywords) opts.keywords = cp.keywords;
      if (cp.description) opts.description = cp.description;
      if (cp.lastModifiedBy) opts.lastModifiedBy = cp.lastModifiedBy;
      if (cp.revision) opts.revision = parseInt(cp.revision, 10);
    }
  }

  // 3. Parse show options from presProps
  if (pptx.presProps) {
    const presPropsEl = pptx.doc.get(pptx.presProps);
    if (presPropsEl) {
      const showPr = findChild(presPropsEl, "p:showPr");
      if (showPr) {
        const show: Record<string, unknown> = {};
        if (attr(showPr, "loop") === "1") show.loop = true;
        // kiosk is a child element per XSD EG_ShowType, not an attribute
        if (findChild(showPr, "p:kiosk")) show.kiosk = true;
        const showNarrationVal = attr(showPr, "showNarration");
        if (showNarrationVal === "0") show.showNarration = false;
        else if (showNarrationVal !== undefined) show.showNarration = true;
        if (attr(showPr, "useTimings") === "1") show.useTimings = true;
        if (Object.keys(show).length > 0) opts.show = show;
      }
    }
  }

  // 4. Build relationship maps
  const masterThemePaths = resolveRelTargets(pptx.doc, pptx.slideMasters, (t) =>
    t.includes("/theme"),
  );
  const layoutMasterPaths = resolveRelTargets(pptx.doc, pptx.slideLayouts, (t) =>
    t.includes("/slideMaster"),
  );
  const slideLayoutPaths = resolveRelTargets(pptx.doc, pptx.slides, (t) =>
    t.includes("/slideLayout"),
  );

  // 5. Parse masters
  const masterCount = pptx.slideMasters.length;
  const masterDefs: MasterDefinition[] = [];
  for (let mi = 0; mi < masterCount; mi++) {
    const masterPath = pptx.slideMasters[mi];
    const masterEl = pptx.doc.get(masterPath);
    if (!masterEl) continue;

    // Theme
    const themePath = masterThemePaths.get(masterPath);
    const themeEl = themePath ? pptx.doc.get(themePath) : undefined;
    const themeOptions = themeEl ? parseTheme(themeEl) : undefined;

    // Background (inline: findChild → parseBackground)
    const cSld = findChild(masterEl, "p:cSld");
    const bg = cSld ? findChild(cSld, "p:bg") : undefined;
    const masterBackground = bg ? parseBackground(bg) : undefined;
    const hasBackground = masterBackground && Object.keys(masterBackground).length > 0;

    // Layouts belonging to this master
    const masterLayouts: NonNullable<MasterDefinition["layouts"]>[number][] = [];
    for (const layoutPath of pptx.slideLayouts) {
      if (layoutMasterPaths.get(layoutPath) !== masterPath) continue;
      const layoutEl = pptx.doc.get(layoutPath);
      if (layoutEl) masterLayouts.push({ type: parseSlideLayoutType(layoutEl) as SlideLayoutType });
    }

    const masterName = themeOptions?.name ?? `master${mi + 1}`;
    const masterDef = {} as Record<string, unknown>;
    masterDef.name = masterName;
    if (themeOptions) masterDef.theme = themeOptions;
    if (hasBackground) masterDef.background = masterBackground;
    if (masterLayouts.length > 0) masterDef.layouts = masterLayouts;
    masterDefs.push(masterDef as MasterDefinition);
  }

  // Only set masters if there's more than one (single master is auto-created)
  if (masterCount > 1) {
    opts.masters = masterDefs;
  }

  // 6. Parse comment authors
  const commentAuthors = new Map<number, { name: string; initials: string }>();
  if (pptx.partRefs.commentAuthors) {
    const authorsEl = pptx.doc.get(pptx.partRefs.commentAuthors);
    if (authorsEl) {
      for (const child of authorsEl.elements ?? []) {
        if (child.name !== "p:cmAuthor") continue;
        const id = attrNum(child, "id");
        const name = attr(child, "name");
        const initials = attr(child, "initials");
        if (id !== undefined) {
          commentAuthors.set(id, { name: name ?? "", initials: initials ?? "" });
        }
      }
    }
  }

  // 7. Parse slides with layout and master references
  const result: SlideOptions[] = [];
  for (let si = 0; si < pptx.slides.length; si++) {
    const slidePath = pptx.slides[si];
    const slideEl = pptx.doc.get(slidePath);
    if (!slideEl) continue;

    const slideRels = parseSlideRelMap(pptx.doc, slidePath);
    const ctx = new ParseContext(pptx, slideRels);
    const slideOpts = parseSlide(slideEl, ctx) as Record<string, unknown>;

    // Resolve layout → master
    const layoutPath = slideLayoutPaths.get(slidePath);
    if (layoutPath) {
      const layoutEl = pptx.doc.get(layoutPath);
      if (layoutEl) slideOpts.layout = parseSlideLayoutType(layoutEl);

      const resolvedMasterPath = layoutMasterPaths.get(layoutPath);
      if (resolvedMasterPath) {
        const masterIdx = pptx.slideMasters.indexOf(resolvedMasterPath);
        if (masterIdx >= 0 && masterDefs[masterIdx]) {
          slideOpts.master = masterDefs[masterIdx].name;
        }
      }
    }

    // Comments via slide rels
    for (const [, relPath] of slideRels) {
      if (!relPath.includes("/comments/")) continue;
      const commentsEl = pptx.doc.get(relPath);
      if (!commentsEl) continue;

      const comments: Record<string, unknown>[] = [];
      for (const cm of commentsEl.elements ?? []) {
        if (cm.name !== "p:cm") continue;
        const authorId = attrNum(cm, "authorId");
        const dt = attr(cm, "dt");
        const idx = attrNum(cm, "idx");

        const author = authorId !== undefined ? commentAuthors.get(authorId) : undefined;
        const pos = findChild(cm, "p:pos");
        const x = pos ? attrNum(pos, "x") : undefined;
        const y = pos ? attrNum(pos, "y") : undefined;

        // Extract text from p:txBody
        const txBody = findChild(cm, "p:txBody");
        const text = txBody ? extractCommentText(txBody) : "";

        const commentEntry: Record<string, unknown> = {};
        if (authorId !== undefined) commentEntry.authorId = authorId;
        if (author) commentEntry.author = author.name;
        if (idx !== undefined) commentEntry.idx = idx;
        if (dt) commentEntry.date = dt;
        if (x !== undefined) commentEntry.x = x;
        if (y !== undefined) commentEntry.y = y;
        if (text) commentEntry.text = text;

        comments.push(commentEntry);
      }

      if (comments.length > 0) slideOpts.comments = comments;
      break;
    }

    result.push(slideOpts as SlideOptions);
  }

  opts.slides = result;
  return opts as PresentationOptions;
}

/**
 * Extract plain text from a p:txBody in a comment element.
 */
function extractCommentText(txBody: Element): string {
  const parts: string[] = [];
  for (const p of txBody.elements ?? []) {
    if (p.name !== "a:p") continue;
    for (const r of p.elements ?? []) {
      if (r.name !== "a:r") continue;
      for (const t of r.elements ?? []) {
        if (t.name === "a:t") {
          const text = textOf(t);
          if (text) parts.push(text);
        }
      }
    }
  }
  return parts.join("");
}
