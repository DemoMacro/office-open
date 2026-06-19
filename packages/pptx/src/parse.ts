import type { ParsedArchive } from "@office-open/core";
import {
  appPropertiesDesc,
  customPropertiesDesc,
  parseArchive,
  parseCorePropsElement,
  convertEmuToPixels,
} from "@office-open/core";
import type { DataType } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import { themeDesc } from "@office-open/core/theme";
import type { Element } from "@office-open/xml";
import { attr, attrNum, findChild } from "@office-open/xml";

import { PptxReadContext, ParseContext } from "./context";
import { backgroundDesc } from "./parts/descriptors/background";
import { parseChild } from "./parts/descriptors/bridge";
import { commentAuthorsDesc, slideCommentsDesc } from "./parts/descriptors/comments";
import { notesMasterDesc } from "./parts/descriptors/notes-master";
import { notesSlideDesc } from "./parts/descriptors/notes-slide";
import { presPropsDesc } from "./parts/descriptors/presentation-properties";
import { slideDesc } from "./parts/descriptors/slide";
import { slideLayoutDesc } from "./parts/descriptors/slide-layout";
import { tableStylesDesc } from "./parts/descriptors/table-styles";
import { viewPropsDesc } from "./parts/descriptors/view-properties";
import type { SlideChild } from "./parts/slide/slide-child";

export { parseArchive };

import type { SlideLayoutType } from "./parts/slide-layout";
import type {
  MasterDefinition,
  SlideOptions,
  SlideCommentOptions,
  PresentationOptions,
} from "./shared/file";

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
  doc: ParsedArchive;
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
  /** docProps/custom.xml */
  customProps?: string;
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

function parseRootRels(doc: ParsedArchive): {
  coreProps?: string;
  appProps?: string;
  customProps?: string;
} {
  const relsEl = doc.get("_rels/.rels");
  if (!relsEl) return {};

  let coreProps: string | undefined;
  let appProps: string | undefined;
  let customProps: string | undefined;

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
    } else if (type.includes("/custom-properties")) {
      customProps = path;
    }
  }

  return { coreProps, appProps, customProps };
}

function parseSlideRels(doc: ParsedArchive, slidePaths: string[], refs: PptxPartRefs): void {
  const commentsSet = new Set(refs.comments);
  const chartsSet = new Set(refs.charts);
  const diagramDataSet = new Set(refs.diagramData);
  const mediaSet = new Set(refs.media);

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
        commentsSet.add(path);
      } else if (type.includes("/chart")) {
        chartsSet.add(path);
      } else if (type.includes("/diagramData")) {
        diagramDataSet.add(path);
      } else if (type.includes("/image") || type.includes("/video") || type.includes("/media")) {
        mediaSet.add(path);
      }
    }
  }

  refs.comments = [...commentsSet];
  refs.charts = [...chartsSet];
  refs.diagramData = [...diagramDataSet];
  refs.media = [...mediaSet];
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

  const { coreProps, appProps, customProps } = parseRootRels(doc);

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
    customProps,
  };
}

/**
 * Parse a single slide's relationship file into a Map<rId, path>.
 */
function parseSlideRelMap(doc: ParsedArchive, slidePath: string): Map<string, string> {
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
    // External links (hyperlinks) keep their original URL target
    if (attr(child, "TargetMode") === "External") {
      rels.set(id, target);
    } else {
      rels.set(id, resolveRelsPath(target));
    }
  }

  return rels;
}

/**
 * Build a map from each path to the rel target matching a predicate.
 */
function resolveRelTargets(
  doc: ParsedArchive,
  paths: string[],
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
 * Parse p14:sectionLst from presentation.xml and map each slide path to its
 * section name. Bridges p14:sldId (by slide id) -> p:sldIdLst (slide id ->
 * rId) -> presentation rels (rId -> path).
 */
function parseSlideSections(
  presentation: Element | undefined,
  doc: ParsedArchive,
): Map<string, string> {
  const pathToSection = new Map<string, string>();
  if (!presentation) return pathToSection;

  const extLst = findChild(presentation, "p:extLst");
  if (!extLst) return pathToSection;

  let sectionLst: Element | undefined;
  for (const ext of extLst.elements ?? []) {
    if (ext.name !== "p:ext") continue;
    if (attr(ext, "uri") !== "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}") continue;
    sectionLst = findChild(ext, "p14:sectionLst");
    if (sectionLst) break;
  }
  if (!sectionLst) return pathToSection;

  // slideId -> sectionName
  const sectionBySlideId = new Map<number, string>();
  for (const section of sectionLst.elements ?? []) {
    if (section.name !== "p14:section") continue;
    const name = attr(section, "name");
    if (!name) continue;
    const sldIdLst = findChild(section, "p14:sldIdLst");
    for (const sldId of sldIdLst?.elements ?? []) {
      if (sldId.name !== "p14:sldId") continue;
      const id = attrNum(sldId, "id");
      if (id !== undefined) sectionBySlideId.set(id, name);
    }
  }
  if (sectionBySlideId.size === 0) return pathToSection;

  // slideId -> rId (from p:sldIdLst)
  const sldIdLst = findChild(presentation, "p:sldIdLst");
  const rIdBySlideId = new Map<number, string>();
  for (const sldId of sldIdLst?.elements ?? []) {
    if (sldId.name !== "p:sldId") continue;
    const id = attrNum(sldId, "id");
    const rId = attr(sldId, "r:id");
    if (id !== undefined && rId) rIdBySlideId.set(id, rId);
  }

  // rId -> path (from presentation.xml.rels)
  const relsEl = doc.get("ppt/_rels/presentation.xml.rels");
  const pathByRId = new Map<string, string>();
  for (const child of relsEl?.elements ?? []) {
    if (child.name !== "Relationship") continue;
    const id = attr(child, "Id");
    const target = attr(child, "Target");
    if (id && target) pathByRId.set(id, resolveRelsPath(target));
  }

  for (const [slideId, name] of sectionBySlideId) {
    const rId = rIdBySlideId.get(slideId);
    if (!rId) continue;
    const path = pathByRId.get(rId);
    if (path) pathToSection.set(path, name);
  }

  return pathToSection;
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
  const sectionBySlidePath = parseSlideSections(pptx.presentation, pptx.doc);

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
      if (cp.revision !== undefined) opts.revision = cp.revision;
      if (cp.lastPrinted) opts.lastPrinted = cp.lastPrinted;
      if (cp.created) opts.created = cp.created;
      if (cp.modified) opts.modified = cp.modified;
    }
  }

  // 2b. Parse extended (app) properties
  if (pptx.appProps) {
    const appPropsEl = pptx.doc.get(pptx.appProps);
    if (appPropsEl) {
      const ap = appPropertiesDesc.parse(appPropsEl, {} as ReadContext);
      if (ap && Object.keys(ap).length > 0) opts.appProperties = ap;
    }
  }

  // 2c. Parse custom properties
  if (pptx.customProps) {
    const customPropsEl = pptx.doc.get(pptx.customProps);
    if (customPropsEl) {
      const cp = customPropertiesDesc.parse(customPropsEl, {} as ReadContext);
      if (cp.properties?.length) opts.customProperties = cp.properties;
    }
  }

  // 3. Parse show options from presProps
  if (pptx.presProps) {
    const presPropsEl = pptx.doc.get(pptx.presProps);
    if (presPropsEl) {
      const presPropsOpts = presPropsDesc.parse(presPropsEl, {} as ReadContext);
      if (presPropsOpts.show) opts.show = presPropsOpts.show;
    }
  }

  // 3b. Parse view properties
  if (pptx.viewProps) {
    const viewPropsEl = pptx.doc.get(pptx.viewProps);
    if (viewPropsEl) {
      const viewOpts = viewPropsDesc.parse(viewPropsEl, {} as ReadContext);
      if (viewOpts.lastView || viewOpts.showComments !== undefined || viewOpts.gridSpacing) {
        opts.view = viewOpts;
      }
    }
  }

  // 3c. Parse table styles
  if (pptx.tableStyles) {
    const tableStylesEl = pptx.doc.get(pptx.tableStyles);
    if (tableStylesEl) {
      const tableStylesResult = tableStylesDesc.parse(tableStylesEl, {} as ReadContext);
      if (tableStylesResult.opts) opts.tableStyles = tableStylesResult.opts;
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
  const masterReadCtx = new PptxReadContext(new ParseContext(pptx, new Map()));
  for (let mi = 0; mi < masterCount; mi++) {
    const masterPath = pptx.slideMasters[mi];
    const masterEl = pptx.doc.get(masterPath);
    if (!masterEl) continue;

    // Theme
    const themePath = masterThemePaths.get(masterPath);
    const themeEl = themePath ? pptx.doc.get(themePath) : undefined;
    const themeOptions = themeEl ? themeDesc.parse(themeEl, masterReadCtx) : undefined;

    // Background (inline: findChild → backgroundDesc.parse)
    const cSld = findChild(masterEl, "p:cSld");
    const bg = cSld ? findChild(cSld, "p:bg") : undefined;
    const masterBackground = bg ? backgroundDesc.parse(bg, masterReadCtx) : undefined;
    const hasBackground = masterBackground && Object.keys(masterBackground).length > 0;

    // Children: extract non-placeholder shapes from spTree
    const spTree = cSld ? findChild(cSld, "p:spTree") : undefined;
    const masterChildren: SlideChild[] = [];
    if (spTree) {
      for (const child of spTree.elements ?? []) {
        // Skip nvGrpSpPr/grpSpPr (tree container structure, not shapes)
        if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
        // Skip placeholder shapes (generated by DefaultSlideMaster template)
        if (child.name === "p:sp") {
          const nvSpPr = findChild(child, "p:nvSpPr");
          const nvPr = nvSpPr ? findChild(nvSpPr, "p:nvPr") : undefined;
          if (nvPr && findChild(nvPr, "p:ph")) continue;
        }
        const parsed = parseChild(child, masterReadCtx);
        if (parsed !== undefined) masterChildren.push(parsed);
      }
    }

    // Layouts belonging to this master
    const masterLayouts: NonNullable<MasterDefinition["layouts"]>[number][] = [];
    for (const layoutPath of pptx.slideLayouts) {
      if (layoutMasterPaths.get(layoutPath) !== masterPath) continue;
      const layoutEl = pptx.doc.get(layoutPath);
      if (layoutEl) {
        const layoutOpts = slideLayoutDesc.parse(layoutEl, masterReadCtx);
        const layoutDef: Record<string, unknown> = {
          type: (layoutOpts.type ?? "blank") as SlideLayoutType,
        };
        if (layoutOpts.placeholders) layoutDef.placeholders = layoutOpts.placeholders;
        masterLayouts.push(layoutDef as NonNullable<MasterDefinition["layouts"]>[number]);
      }
    }

    const masterName = themeOptions?.name ?? `master${mi + 1}`;
    const masterDef = {} as Record<string, unknown>;
    masterDef.name = masterName;
    if (themeOptions) masterDef.theme = themeOptions;
    if (hasBackground) masterDef.background = masterBackground;
    if (masterChildren.length > 0) masterDef.children = masterChildren;
    if (masterLayouts.length > 0) masterDef.layouts = masterLayouts;
    masterDefs.push(masterDef as MasterDefinition);
  }

  // Only set masters if there's more than one (single master is auto-created)
  if (masterCount > 1) {
    opts.masters = masterDefs;
  }

  // 5b. Parse notes masters
  for (const nmPath of pptx.partRefs.notesMasters) {
    const nmEl = pptx.doc.get(nmPath);
    if (nmEl) {
      const nmOpts = notesMasterDesc.parse(nmEl, masterReadCtx);
      if (nmOpts.options) opts.includeNotesMaster = true;
      if (nmOpts.options) Object.assign(opts, { notesMasterOptions: nmOpts.options });
    }
  }

  // 5c. Parse handout masters (no separate handling needed, just mark inclusion)
  // Handout masters are referenced from presentation.xml rels but not stored in PptxDocument currently

  // 6. Parse comment authors
  const commentAuthors = new Map<number, { name: string; initials: string }>();
  if (pptx.partRefs.commentAuthors) {
    const authorsEl = pptx.doc.get(pptx.partRefs.commentAuthors);
    if (authorsEl) {
      const authors = commentAuthorsDesc.parse(authorsEl, masterReadCtx);
      for (const a of authors) {
        commentAuthors.set(a.id, { name: a.name, initials: a.initials });
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
    const readCtx = new PptxReadContext(ctx);
    const slideOpts = slideDesc.parse(slideEl, readCtx) as Record<string, unknown>;

    // Resolve layout → master
    const layoutPath = slideLayoutPaths.get(slidePath);
    if (layoutPath) {
      const layoutEl = pptx.doc.get(layoutPath);
      if (layoutEl) {
        const layoutOpts = slideLayoutDesc.parse(layoutEl, readCtx);
        slideOpts.layout = (layoutOpts.type ?? "blank") as SlideLayoutType;
      }

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

      const parsedComments = slideCommentsDesc.parse(commentsEl, readCtx);
      if (parsedComments.length > 0) {
        const comments: Partial<SlideCommentOptions>[] = [];
        for (const cm of parsedComments) {
          const entry: Partial<SlideCommentOptions> = { x: cm.x, y: cm.y };
          if (cm.text) entry.text = cm.text;
          if (cm.date) entry.date = cm.date;
          if (cm.modified !== undefined) entry.modified = cm.modified;
          const author = commentAuthors.get(cm.authorId);
          if (author) {
            entry.author = author.name;
            if (author.initials) entry.initials = author.initials;
          }
          comments.push(entry);
        }
        if (comments.length > 0) slideOpts.comments = comments as SlideCommentOptions[];
      }
      break;
    }

    // Notes slide via slide rels
    for (const [, relPath] of slideRels) {
      if (!relPath.includes("/notesSlides/")) continue;
      const notesEl = pptx.doc.get(relPath);
      if (!notesEl) continue;
      const notesData = notesSlideDesc.parse(notesEl, readCtx);
      if (notesData.text) slideOpts.notes = notesData.text;
      break;
    }

    // Section (p14:sectionLst) — bridged via slide id -> rId -> path
    const sectionName = sectionBySlidePath.get(slidePath);
    if (sectionName) slideOpts.section = sectionName;

    result.push(slideOpts as SlideOptions);
  }

  opts.slides = result;
  return opts as PresentationOptions;
}
