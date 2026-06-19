/**
 * Descriptor-based PPTX compiler.
 *
 * Produces a valid PPTX ZIP archive using the descriptor pipeline.
 * Accepts pure JSON PresentationOptions — no intermediate class needed.
 *
 * @module
 */

import { Relationships, convertPixelsToEmu } from "@office-open/core";
import type { RelationshipType } from "@office-open/core";
import {
  appPropertiesDesc,
  buildCorePropertiesXmlString,
  collectPlaceholderKeys,
  compileMapping,
  customPropertiesDesc,
  getReferencedMedia,
  getMediaRefs,
  getVideoRefs,
  hasPlaceholders,
  levelForMediaName,
  replaceChartPlaceholders,
  replaceHyperlinkPlaceholders,
  replaceImagePlaceholders,
  replaceMediaPlaceholders,
  replaceSmartArtPlaceholders,
  replaceVideoPlaceholders,
  addSmartArtRelationships,
} from "@office-open/core";
import type { XmlifyedFile, ZipOptions, Zippable } from "@office-open/core";
import { ChartCollection } from "@office-open/core/chart";
import { SmartArtCollection } from "@office-open/core/smartart";
import type { AuthorEntry, CommentEntry } from "@parts/comment";
import { ContentTypes } from "@parts/content-types";
import type { PresentationPartOptions, PresentationSectionGroup } from "@parts/presentation";
import { buildCustomLayoutXml, buildLayoutXml, type SlideLayoutType } from "@parts/slide-layout";
import { buildSlideMasterXml } from "@parts/slide-master";
import type { SlideSyncOptions } from "@parts/slide/slide-sync-properties";
import { getColorXml, getLayoutXml, getStyleXml, DEFAULT_DRAWING_XML } from "@parts/smartart";
import { SP_TREE_HEADER } from "@shared/constants";
import {
  type PresentationOptions,
  type MasterDefinition,
  type SlideOptions,
  type SlideSize,
} from "@shared/file";
import { HyperlinkCollection } from "@shared/hyperlink-collection";
import type { MediaData } from "@shared/media/data";
import { Media } from "@shared/media/media";
import { createThemeXml } from "@shared/theme";
import { buildTransition } from "@shared/transition";

import { PptxWriteContext } from "./context";
import { timingDesc } from "./parts/descriptors/animation";
import { backgroundDesc } from "./parts/descriptors/background";
import { stringifyChild } from "./parts/descriptors/bridge";
import { commentAuthorsDesc, slideCommentsDesc } from "./parts/descriptors/comments";
import { contentTypesDesc } from "./parts/descriptors/content-types";
import { handoutMasterDesc } from "./parts/descriptors/handout-master";
import { notesMasterDesc } from "./parts/descriptors/notes-master";
import { notesSlideDesc } from "./parts/descriptors/notes-slide";
import { presentationDesc } from "./parts/descriptors/presentation";
import { presPropsDesc } from "./parts/descriptors/presentation-properties";
import { slideLayoutDesc } from "./parts/descriptors/slide-layout";
import { slideMasterDesc } from "./parts/descriptors/slide-master";
import { slideSyncDesc } from "./parts/descriptors/slide-sync";
import { tableStylesDesc } from "./parts/descriptors/table-styles";
import { viewPropsDesc } from "./parts/descriptors/view-properties";

// ── Constants ──

const encoder = new TextEncoder();
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// ── Helper types ──

interface RelEntry {
  id: number | string;
  type: RelationshipType;
  target: string;
  mode?: string;
}

interface LayoutInfo {
  key: string;
  index: number;
  masterIndex: number;
  layout: string;
}

interface MasterInfo {
  name: string;
  index: number;
  master: string;
  theme: string;
  layouts: LayoutInfo[];
  masterRels: Relationships;
  layoutRels: Relationships[];
}

interface XmlifyedFileMapping {
  [key: string]: { data: string; path: string };
}

// ── Pure helper functions (extracted from File class) ──

function buildRels(entries: RelEntry[]): Relationships {
  const rels = new Relationships();
  for (const e of entries) {
    rels.addRelationship(e.id, e.type, e.target, e.mode as "External" | undefined);
  }
  return rels;
}

function resolveSlideSize(size?: SlideSize): { width: number; height: number } {
  if (!size || size === "16:9") return { width: 12192000, height: 6858000 };
  if (size === "4:3") return { width: 9144000, height: 6858000 };
  return { width: convertPixelsToEmu(size.width), height: convertPixelsToEmu(size.height) };
}

function deriveInitials(name: string): string {
  const parts = name.trim().split(/\s+/);
  return parts.length >= 2
    ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
    : name.slice(0, 2).toUpperCase();
}

function buildMasterMap(
  masterDefs: MasterDefinition[],
  slides: SlideOptions[],
  slideWidth: number,
): MasterInfo[] {
  const defs = masterDefs.length > 0 ? masterDefs : [{} as MasterDefinition];
  const slideMasterLookup = new Map<number, number>();

  for (let si = 0; si < slides.length; si++) {
    const masterName = slides[si].master;
    if (masterName === undefined) {
      slideMasterLookup.set(si, 0);
      continue;
    }
    const mi = defs.findIndex((d) => d.name === masterName);
    slideMasterLookup.set(si, mi >= 0 ? mi : 0);
  }

  let globalLayoutIndex = 0;
  const masters: MasterInfo[] = [];

  for (let mi = 0; mi < defs.length; mi++) {
    const def = defs[mi];
    const name = def.name ?? `master${mi + 1}`;

    const layoutDefs = def.layouts;
    let layoutKeys: string[];
    if (layoutDefs && layoutDefs.length > 0) {
      layoutKeys = layoutDefs.map(
        (ld) => ld.type ?? ld.name ?? `layout${mi}_${layoutDefs.indexOf(ld)}`,
      );
    } else {
      const seen = new Set<string>();
      const keys: string[] = [];
      for (let si = 0; si < slides.length; si++) {
        if (slideMasterLookup.get(si) === mi) {
          const lt = slides[si].layout ?? "blank";
          if (!seen.has(lt)) {
            seen.add(lt);
            keys.push(lt);
          }
        }
      }
      layoutKeys = keys.length > 0 ? keys : ["blank"];
    }

    const hf = slides.find((s) => s.headerFooter)?.headerFooter;
    const master = buildSlideMasterXml(layoutKeys.length, hf, def, slideWidth, mi);
    const theme = createThemeXml(def.theme);

    const layouts: LayoutInfo[] = [];
    const layoutRels: Relationships[] = [];

    for (let li = 0; li < layoutKeys.length; li++) {
      const key = layoutKeys[li];
      const layoutDef = layoutDefs?.[li];
      const slideLayoutType = (layoutDef?.type ?? key) as SlideLayoutType;
      layouts.push({
        key,
        index: globalLayoutIndex,
        masterIndex: mi,
        layout: layoutDef
          ? buildCustomLayoutXml(layoutDef)
          : buildLayoutXml(slideLayoutType, slideWidth),
      });
      layoutRels.push(
        buildRels([
          {
            id: 1,
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
            target: `../slideMasters/slideMaster${mi + 1}.xml`,
          },
        ]),
      );
      globalLayoutIndex++;
    }

    const masterRelsEntries: RelEntry[] = [];
    for (let li = 0; li < layouts.length; li++) {
      masterRelsEntries.push({
        id: li + 1,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        target: `../slideLayouts/slideLayout${layouts[li].index + 1}.xml`,
      });
    }
    masterRelsEntries.push({
      id: layouts.length + 1,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      target: `../theme/theme${mi + 1}.xml`,
    });

    masters.push({
      name,
      index: mi,
      master,
      theme,
      layouts,
      masterRels: buildRels(masterRelsEntries),
      layoutRels,
    });
  }

  return masters;
}

function findLayoutForSlide(
  masters: MasterInfo[],
  slides: SlideOptions[],
  slideIndex: number,
): LayoutInfo {
  const opts = slides[slideIndex];
  const mi =
    opts.master !== undefined
      ? Math.max(
          0,
          masters.findIndex((m) => m.name === opts.master),
        )
      : 0;
  const master = masters[mi];
  const layoutKey = opts.layout ?? "blank";
  const li = master.layouts.find((l) => l.key === layoutKey);
  return li ?? master.layouts[0];
}

function buildSlideRels(masters: MasterInfo[], slides: SlideOptions[]): Relationships[] {
  const rels: Relationships[] = [];
  for (let i = 0; i < slides.length; i++) {
    const layout = findLayoutForSlide(masters, slides, i);
    rels.push(
      buildRels([
        {
          id: 1,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
          target: `../slideLayouts/slideLayout${layout.index + 1}.xml`,
        },
      ]),
    );
  }
  return rels;
}

function buildCommentData(slides: SlideOptions[]): {
  authors: AuthorEntry[] | undefined;
  perSlide: (CommentEntry[] | undefined)[];
} {
  const authorMap = new Map<
    string,
    { id: number; name: string; initials: string; clrIdx: number; commentCount: number }
  >();
  let nextAuthorId = 0;

  const perSlide: (CommentEntry[] | undefined)[] = Array.from({ length: slides.length });

  for (let i = 0; i < slides.length; i++) {
    const slideComments = slides[i].comments;
    if (!slideComments || slideComments.length === 0) continue;

    const commentEntries: CommentEntry[] = [];

    for (const c of slideComments) {
      let author = authorMap.get(c.author);
      if (!author) {
        const id = nextAuthorId++;
        author = {
          id,
          name: c.author,
          initials: c.initials || deriveInitials(c.author),
          clrIdx: id,
          commentCount: 0,
        };
        authorMap.set(c.author, author);
      }
      author.commentCount++;

      commentEntries.push({
        authorId: author.id,
        idx: author.commentCount,
        date: c.date,
        x: c.x,
        y: c.y,
        text: c.text,
        modified: c.modified,
      });
    }

    perSlide[i] = commentEntries;
  }

  const authors =
    authorMap.size > 0
      ? Array.from(authorMap.values(), (a) => ({
          id: a.id,
          name: a.name,
          initials: a.initials,
          clrIdx: a.clrIdx,
          lastIdx: a.commentCount,
        }))
      : undefined;

  return { authors, perSlide };
}

function initContentTypes(slides: SlideOptions[], includeHandout: boolean): ContentTypes {
  const ct = new ContentTypes();
  let hasComments = false;
  let notesSlideIdx = 0;
  let slideSyncIdx = 0;
  for (let i = 0; i < slides.length; i++) {
    ct.addSlide(i + 1);
    if (slides[i].notes) {
      ct.addNotesSlide(notesSlideIdx + 1);
      notesSlideIdx++;
    }
    if (slides[i].comments && slides[i].comments!.length > 0) {
      ct.addComments(i + 1);
      hasComments = true;
    }
    if (slides[i].slideSync) {
      ct.addSlideSyncPr(slideSyncIdx + 1);
      slideSyncIdx++;
    }
  }
  if (notesSlideIdx > 0) ct.addNotesMaster();
  if (hasComments) ct.addCommentAuthors();
  if (includeHandout) ct.addHandoutMaster();
  return ct;
}

function buildFileRels(hasCustomProperties: boolean): Relationships {
  const entries: RelEntry[] = [
    {
      id: 1,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
      target: "ppt/presentation.xml",
    },
    {
      id: 2,
      type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
      target: "docProps/core.xml",
    },
    {
      id: 3,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
      target: "docProps/app.xml",
    },
  ];
  if (hasCustomProperties) {
    entries.push({
      id: 4,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
      target: "docProps/custom.xml",
    });
  }
  return buildRels(entries);
}

function initPresRels(masters: MasterInfo[], slideCount: number): Relationships {
  const rels = new Relationships();
  let rid = 1;
  for (let mi = 0; mi < masters.length; mi++) {
    rels.addRelationship(
      rid++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
      `slideMasters/slideMaster${mi + 1}.xml`,
    );
  }
  for (let i = 0; i < slideCount; i++) {
    rels.addRelationship(
      rid++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
      `slides/slide${i + 1}.xml`,
    );
  }
  rels.addRelationship(
    rid++,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
    "presProps.xml",
  );
  rels.addRelationship(
    rid++,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
    "viewProps.xml",
  );
  for (let mi = 0; mi < masters.length; mi++) {
    rels.addRelationship(
      rid++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      `theme/theme${mi + 1}.xml`,
    );
  }
  rels.addRelationship(
    rid,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
    "tableStyles.xml",
  );
  return rels;
}

function buildPresAttrOpts(
  options: PresentationOptions,
): Partial<
  Pick<
    PresentationPartOptions,
    | "serverZoom"
    | "firstSlideNum"
    | "showSpecialPlsOnTitleSld"
    | "rtl"
    | "removePersonalInfoOnSave"
    | "compatMode"
    | "strictFirstAndLastChars"
    | "embedTrueTypeFonts"
    | "saveSubsetFonts"
    | "autoCompressPictures"
    | "bookmarkIdSeed"
    | "conformance"
    | "photoAlbum"
    | "modifyVerifier"
    | "embeddedFonts"
    | "customShows"
    | "kinsoku"
    | "customerData"
  >
> {
  if (
    !options.serverZoom &&
    options.firstSlideNum === undefined &&
    options.showSpecialPlsOnTitleSld === undefined &&
    options.rtl === undefined &&
    options.removePersonalInfoOnSave === undefined &&
    options.compatMode === undefined &&
    options.strictFirstAndLastChars === undefined &&
    options.embedTrueTypeFonts === undefined &&
    options.saveSubsetFonts === undefined &&
    options.autoCompressPictures === undefined &&
    options.bookmarkIdSeed === undefined &&
    options.conformance === undefined &&
    options.photoAlbum === undefined &&
    options.modifyVerifier === undefined &&
    options.embeddedFonts === undefined &&
    options.customShows === undefined &&
    options.kinsoku === undefined &&
    options.customerData === undefined
  ) {
    return {};
  }
  return {
    serverZoom: options.serverZoom,
    firstSlideNum: options.firstSlideNum,
    showSpecialPlsOnTitleSld: options.showSpecialPlsOnTitleSld,
    rtl: options.rtl,
    removePersonalInfoOnSave: options.removePersonalInfoOnSave,
    compatMode: options.compatMode,
    strictFirstAndLastChars: options.strictFirstAndLastChars,
    embedTrueTypeFonts: options.embedTrueTypeFonts,
    saveSubsetFonts: options.saveSubsetFonts,
    autoCompressPictures: options.autoCompressPictures,
    bookmarkIdSeed: options.bookmarkIdSeed,
    conformance: options.conformance,
    photoAlbum: options.photoAlbum,
    modifyVerifier: options.modifyVerifier,
    embeddedFonts: options.embeddedFonts,
    customShows: options.customShows,
    kinsoku: options.kinsoku,
    customerData: options.customerData,
  };
}

// ── Slide serializer using descriptors ──

/**
 * Serialize a single slide to its `<p:sld>` XML (no XML declaration — matches
 * the generated slide parts). Exposed so patch can append/replace slides by
 * reusing the full slide vocabulary without re-running the compiler.
 */
export function stringifySlide(slideOpts: SlideOptions, ctx: PptxWriteContext): string {
  const parts: string[] = [];

  const sldAttrs: string[] = [];
  if (slideOpts.showMasterShapes === false) sldAttrs.push(' showMasterSp="0"');
  if (slideOpts.showMasterPlaceholderAnimations === false) sldAttrs.push(' showMasterPhAnim="0"');
  parts.push(
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${sldAttrs.join("")}>`,
  );

  parts.push("<p:cSld>");

  if (slideOpts.background) {
    parts.push(backgroundDesc.stringify(slideOpts.background, ctx) ?? "");
  }

  parts.push("<p:spTree>");
  parts.push(SP_TREE_HEADER);

  if (slideOpts.children) {
    for (const child of slideOpts.children) {
      const xml = stringifyChild(child, ctx);
      if (xml) parts.push(xml);
    }
  }

  parts.push("</p:spTree>");

  if (slideOpts.customerData && slideOpts.customerData.length > 0) {
    const cdItems = slideOpts.customerData.map((d) => `<p:custData r:id="${d.rId}"/>`).join("");
    parts.push(`<p:custDataLst>${cdItems}</p:custDataLst>`);
  }

  if (slideOpts.controls && slideOpts.controls.length > 0) {
    const ctrlItems = slideOpts.controls
      .map((c) => {
        const attrs: string[] = [];
        if (c.shapeId !== undefined) attrs.push(`spid="${c.shapeId}"`);
        if (c.name) attrs.push(`name="${c.name}"`);
        if (c.showAsIcon) attrs.push('showAsIcon="1"');
        if (c.rId) attrs.push(`r:id="${c.rId}"`);
        if (c.imageWidth !== undefined) attrs.push(`imgW="${c.imageWidth}"`);
        if (c.imageHeight !== undefined) attrs.push(`imgH="${c.imageHeight}"`);
        return `<p:control ${attrs.join(" ")}/>`;
      })
      .join("");
    parts.push(`<p:controls>${ctrlItems}</p:controls>`);
  }

  parts.push("</p:cSld>");
  parts.push("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");

  if (slideOpts.transition) {
    parts.push(buildTransition(slideOpts.transition));
  }

  if (slideOpts.animations && slideOpts.animations.length > 0) {
    const entries = slideOpts.animations.map((a) => ({ spid: a.shapeId, options: a.options }));
    parts.push(timingDesc.stringify({ entries }, ctx) ?? "");
  }

  parts.push("</p:sld>");
  return parts.join("");
}

// ── Main compiler entry ──

export function compilePresentation(
  options: PresentationOptions,
  overrides: XmlifyedFile[] = [],
  mediaLevel: number = 0,
): Zippable {
  const descCtx = new PptxWriteContext();
  const slides = options.slides ?? [];
  const masterDefs = options.masters ?? [];
  const sz = resolveSlideSize(options.size);
  const includeHandout = options.includeHandoutMaster ?? false;

  // ── Pure structural computations ──

  const masters = buildMasterMap(masterDefs, slides, sz.width);
  const allLayouts = masters.flatMap((m) => m.layouts);
  const allLayoutRels = masters.flatMap((m) => m.layoutRels);
  const themes = masters.map((m) => m.theme);
  const masterRels = masters.map((m) => m.masterRels);
  const slideRels = buildSlideRels(masters, slides);
  const { authors: commentAuthorEntries, perSlide: slideCommentEntries } = buildCommentData(slides);

  const notesTexts: string[] = [];
  const notesSlideIndexMap = new Map<number, number>();
  let notesIdx = 0;
  for (let i = 0; i < slides.length; i++) {
    if (slides[i].notes) {
      notesTexts.push(slides[i].notes!);
      notesSlideIndexMap.set(i, notesIdx++);
    }
  }

  const slideSyncOptionsList: SlideSyncOptions[] = [];
  const slideSyncIndexMap = new Map<number, number>();
  let syncIdx = 0;
  for (let i = 0; i < slides.length; i++) {
    if (slides[i].slideSync) {
      slideSyncOptionsList.push(slides[i].slideSync!);
      slideSyncIndexMap.set(i, syncIdx++);
    }
  }

  // ── Mutable state ──

  const hasCustomProperties = !!options.customProperties && options.customProperties.length > 0;
  const contentTypes = initContentTypes(slides, includeHandout);
  if (hasCustomProperties) contentTypes.addCustomProperties();
  const presRels = initPresRels(masters, slides.length);
  // Group slides into p14:sections by name (first-occurrence order); slides
  // without a section name are left ungrouped (absent from p14:sectionLst).
  const sectionOrder: string[] = [];
  const sectionIndices = new Map<string, number[]>();
  for (let i = 0; i < slides.length; i++) {
    const name = slides[i].section;
    if (!name) continue;
    let arr = sectionIndices.get(name);
    if (!arr) {
      arr = [];
      sectionIndices.set(name, arr);
      sectionOrder.push(name);
    }
    arr.push(i);
  }
  const sections: PresentationSectionGroup[] = sectionOrder.map((name) => ({
    name,
    slideIndices: sectionIndices.get(name)!,
  }));

  const presOptions: PresentationPartOptions = {
    slideWidth: sz.width,
    slideHeight: sz.height,
    slideIds: slides.map((_, i) => 256 + i),
    masterCount: masters.length,
    sections,
    ...buildPresAttrOpts(options),
  };
  const fileRels = buildFileRels(hasCustomProperties);
  const media = new Media();
  const charts = new ChartCollection();
  const smartArts = new SmartArtCollection();
  const hyperlinks = new HyperlinkCollection();

  const presPropsFullOpts =
    options.web || options.print || options.htmlPublish || options.colorMru
      ? {
          web: options.web,
          print: options.print,
          htmlPublish: options.htmlPublish,
          colorMru: options.colorMru,
        }
      : undefined;

  const hasOutlineViewSlides =
    !!options.view?.outlineView?.slides && options.view.outlineView.slides.length > 0;
  const htmlPublishInfo = presPropsFullOpts?.htmlPublish?.rId
    ? { rId: presPropsFullOpts.htmlPublish.rId, target: presPropsFullOpts.htmlPublish.target }
    : undefined;

  // ── Build XML file mapping ──

  const mapping: XmlifyedFileMapping = {
    AppProperties: {
      data: XML_DECL + (appPropertiesDesc.stringify(options.appProperties ?? {}, descCtx) ?? ""),
      path: "docProps/app.xml",
    },
    Properties: {
      data: XML_DECL + buildCorePropertiesXmlString(options),
      path: "docProps/core.xml",
    },
    ...(hasCustomProperties
      ? {
          CustomProperties: {
            data:
              XML_DECL +
              (customPropertiesDesc.stringify(
                { properties: options.customProperties ?? [] },
                descCtx,
              ) ?? ""),
            path: "docProps/custom.xml",
          },
        }
      : {}),
    FileRelationships: {
      data: XML_DECL + fileRels.serialize(),
      path: "_rels/.rels",
    },
  };

  for (let ti = 0; ti < themes.length; ti++) {
    mapping[`Theme${ti}`] = {
      data: XML_DECL + themes[ti],
      path: `ppt/theme/theme${ti + 1}.xml`,
    };
  }

  mapping["TableStyles"] = {
    data: XML_DECL + (tableStylesDesc.stringify({ opts: options.tableStyles }, descCtx) ?? ""),
    path: "ppt/tableStyles.xml",
  };

  mapping["PresProps"] = {
    data: XML_DECL + (presPropsDesc.stringify(presPropsFullOpts ?? {}, descCtx) ?? ""),
    path: "ppt/presProps.xml",
  };

  mapping["ViewProps"] = {
    data: XML_DECL + (viewPropsDesc.stringify(options.view ?? {}, descCtx) ?? ""),
    path: "ppt/viewProps.xml",
  };

  // Slide Masters
  for (let mi = 0; mi < masters.length; mi++) {
    mapping[`SlideMaster${mi}`] = {
      data: XML_DECL + (slideMasterDesc.stringify({ master: masters[mi].master }, descCtx) ?? ""),
      path: `ppt/slideMasters/slideMaster${mi + 1}.xml`,
    };
    mapping[`SlideMasterRels${mi}`] = {
      data: XML_DECL + masterRels[mi].serialize(),
      path: `ppt/slideMasters/_rels/slideMaster${mi + 1}.xml.rels`,
    };
  }

  // Slide Layouts
  for (let li = 0; li < allLayouts.length; li++) {
    mapping[`SlideLayout${li}`] = {
      data:
        XML_DECL + (slideLayoutDesc.stringify({ layout: allLayouts[li].layout }, descCtx) ?? ""),
      path: `ppt/slideLayouts/slideLayout${li + 1}.xml`,
    };
    mapping[`SlideLayoutRels${li}`] = {
      data: XML_DECL + allLayoutRels[li].serialize(),
      path: `ppt/slideLayouts/_rels/slideLayout${li + 1}.xml.rels`,
    };
  }

  // Register content types
  for (let i = 0; i < allLayouts.length; i++) contentTypes.addSlideLayout(i + 1);
  for (let mi = 0; mi < masters.length; mi++) {
    contentTypes.addSlideMaster(mi + 1);
    contentTypes.addTheme(mi + 1);
  }

  // Notes Master
  if (notesTexts.length > 0) {
    const notesMasterRId = presRels.relationshipCount + 1;
    presRels.addRelationship(
      notesMasterRId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
      "notesMasters/notesMaster1.xml",
    );
    presOptions.notesMasterRId = notesMasterRId;
    const notesMasterThemeIndex = themes.length + 1;
    mapping["NotesMaster"] = {
      data:
        XML_DECL +
        (notesMasterDesc.stringify({ options: options.notesMasterOptions }, descCtx) ?? ""),
      path: "ppt/notesMasters/notesMaster1.xml",
    };
    const notesMasterThemeXml = createThemeXml();
    mapping["NotesMasterTheme"] = {
      data: XML_DECL + notesMasterThemeXml,
      path: `ppt/theme/theme${notesMasterThemeIndex}.xml`,
    };
    contentTypes.addTheme(notesMasterThemeIndex);
    const notesMasterRels = new Relationships();
    notesMasterRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      `../theme/theme${notesMasterThemeIndex}.xml`,
    );
    mapping["NotesMasterRelationships"] = {
      data: XML_DECL + notesMasterRels.serialize(),
      path: "ppt/notesMasters/_rels/notesMaster1.xml.rels",
    };
  }

  // Handout Master
  if (includeHandout) {
    const handoutMasterRId = presRels.relationshipCount + 1;
    presRels.addRelationship(
      handoutMasterRId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster",
      "handoutMasters/handoutMaster1.xml",
    );
    presOptions.handoutMasterRId = handoutMasterRId;
    const handoutMasterThemeIndex = themes.length + (notesTexts.length > 0 ? 2 : 1);
    mapping["HandoutMaster"] = {
      data:
        XML_DECL +
        (handoutMasterDesc.stringify({ options: options.handoutMasterOptions }, descCtx) ?? ""),
      path: "ppt/handoutMasters/handoutMaster1.xml",
    };
    const handoutMasterThemeXml = createThemeXml();
    mapping["HandoutMasterTheme"] = {
      data: XML_DECL + handoutMasterThemeXml,
      path: `ppt/theme/theme${handoutMasterThemeIndex}.xml`,
    };
    contentTypes.addTheme(handoutMasterThemeIndex);
    const handoutMasterRels = new Relationships();
    handoutMasterRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      `../theme/theme${handoutMasterThemeIndex}.xml`,
    );
    mapping["HandoutMasterRelationships"] = {
      data: XML_DECL + handoutMasterRels.serialize(),
      path: "ppt/handoutMasters/_rels/handoutMaster1.xml.rels",
    };
  }

  // Comment Authors
  if (commentAuthorEntries) {
    presRels.addRelationship(
      presRels.relationshipCount + 1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
      "commentAuthors.xml",
    );
  }

  // Presentation XML
  const presBody = presentationDesc.stringify(presOptions, descCtx);
  const presentationXml = presBody ? XML_DECL + presBody : "";
  const mediaData = getReferencedMedia(presentationXml, media.array);
  const presImageOffset = presRels.relationshipCount + 1;
  for (let idx = 0; idx < mediaData.length; idx++) {
    presRels.addRelationship(
      presImageOffset + idx,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      `../media/${mediaData[idx].fileName}`,
    );
  }
  const replacedPresentationXml = replaceImagePlaceholders(
    presentationXml,
    mediaData,
    presImageOffset,
  );

  mapping["Presentation"] = {
    data: replacedPresentationXml,
    path: "ppt/presentation.xml",
  };
  mapping["PresentationRelationships"] = {
    data: XML_DECL + presRels.serialize(),
    path: "ppt/_rels/presentation.xml.rels",
  };

  // Slides
  for (let i = 0; i < slides.length; i++) {
    const slideXml = stringifySlide(slides[i], descCtx);

    const slideMediaData = getReferencedMedia(slideXml, media.array);
    const currentSlideRels = slideRels[i];
    const slideImageOffset = currentSlideRels.relationshipCount + 1;
    for (let idx = 0; idx < slideMediaData.length; idx++) {
      currentSlideRels.addRelationship(
        slideImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${slideMediaData[idx].fileName}`,
      );
    }

    let replacedSlideXml = replaceImagePlaceholders(slideXml, slideMediaData, slideImageOffset);

    if (hasPlaceholders(replacedSlideXml)) {
      // Chart
      const slideChartKeys = collectPlaceholderKeys(replacedSlideXml, "chart:");
      if (slideChartKeys.length > 0) {
        const slideChartOffset = currentSlideRels.relationshipCount + 1;
        const slideChartKeySet = new Set(slideChartKeys);
        const xmlCompCharts = charts.array.filter((c) => slideChartKeySet.has(c.key));
        const descCharts = descCtx.charts.filter((c) => slideChartKeySet.has(c.key));
        const allChartKeys = [...xmlCompCharts.map((c) => c.key), ...descCharts.map((c) => c.key)];
        replacedSlideXml = replaceChartPlaceholders(
          replacedSlideXml,
          allChartKeys,
          slideChartOffset,
        );
        for (let ci = 0; ci < allChartKeys.length; ci++) {
          currentSlideRels.addRelationship(
            slideChartOffset + ci,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            `../charts/chart${getChartGlobalIndex(allChartKeys[ci], charts.array, descCtx.charts) + 1}.xml`,
          );
        }
      }

      // SmartArt
      const slideSmartArtKeys = collectPlaceholderKeys(replacedSlideXml, "smartart:");
      if (slideSmartArtKeys.length > 0) {
        const slideSmartArtKeySet = new Set(slideSmartArtKeys);
        const xmlCompSmartArts = smartArts.array.filter((s) => slideSmartArtKeySet.has(s.key));
        const descSmartArts = descCtx.smartArts.filter((s) => slideSmartArtKeySet.has(s.key));
        const allSaKeys = [
          ...xmlCompSmartArts.map((s) => s.key),
          ...descSmartArts.map((s) => s.key),
        ];
        const saOffset = currentSlideRels.relationshipCount + 1;
        replacedSlideXml = replaceSmartArtPlaceholders(replacedSlideXml, allSaKeys, saOffset);
        const saGlobalStart = computeSmartArtGlobalStart(
          allSaKeys[0],
          smartArts.array,
          descCtx.smartArts,
        );
        addSmartArtRelationships(
          allSaKeys,
          (id, type, target) => {
            currentSlideRels.addRelationship(id, type, target);
          },
          saOffset,
          saGlobalStart,
          {
            pathPrefix: "../",
            styleRelType:
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
          },
        );
      }

      // Hyperlinks
      const slideHlinkKeys = collectPlaceholderKeys(replacedSlideXml, "hlink:");
      if (slideHlinkKeys.length > 0) {
        const slideHlinkKeySet = new Set(slideHlinkKeys);
        const slideHlinks = hyperlinks.array.filter((h) => slideHlinkKeySet.has(h.key));
        const hlinkOffset = currentSlideRels.relationshipCount + 1;
        replacedSlideXml = replaceHyperlinkPlaceholders(replacedSlideXml, slideHlinks, hlinkOffset);
        for (let hi = 0; hi < slideHlinks.length; hi++) {
          currentSlideRels.addRelationship(
            hlinkOffset + hi,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            slideHlinks[hi].url,
            "External",
          );
        }
      }

      // Media (video/audio)
      const slideMediaRefs = getMediaRefs(replacedSlideXml, media.array);
      const slideVideoRefs = getVideoRefs(replacedSlideXml, media.array);
      if (slideMediaRefs.length > 0 || slideVideoRefs.length > 0) {
        const mediaOffset = currentSlideRels.relationshipCount + 1;
        const videoOffset = mediaOffset + slideMediaRefs.length;
        replacedSlideXml = replaceMediaPlaceholders(replacedSlideXml, slideMediaRefs, mediaOffset);
        replacedSlideXml = replaceVideoPlaceholders(replacedSlideXml, slideVideoRefs, videoOffset);
        for (let mi = 0; mi < slideMediaRefs.length; mi++) {
          currentSlideRels.addRelationship(
            mediaOffset + mi,
            "http://schemas.microsoft.com/office/2007/relationships/media",
            `../media/${slideMediaRefs[mi].fileName}`,
          );
        }
        for (let vi = 0; vi < slideVideoRefs.length; vi++) {
          currentSlideRels.addRelationship(
            videoOffset + vi,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
            `../media/${slideVideoRefs[vi].fileName}`,
          );
        }
      }
    }

    mapping[`Slide${i}`] = {
      data: replacedSlideXml,
      path: `ppt/slides/slide${i + 1}.xml`,
    };

    if (slideCommentEntries[i]) {
      currentSlideRels.addRelationship(
        currentSlideRels.relationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        `../comments/comment${i + 1}.xml`,
      );
    }

    const notesSlideIndex = notesSlideIndexMap.get(i);
    if (notesSlideIndex !== undefined) {
      currentSlideRels.addRelationship(
        currentSlideRels.relationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
        `../notesSlides/notesSlide${notesSlideIndex + 1}.xml`,
      );
    }

    const slideSyncIndex = slideSyncIndexMap.get(i);
    if (slideSyncIndex !== undefined) {
      currentSlideRels.addRelationship(
        currentSlideRels.relationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideSyncProperties",
        `../slideSyncPr/slideSyncPr${slideSyncIndex + 1}.xml`,
      );
    }

    mapping[`SlideRelationships${i}`] = {
      data: XML_DECL + currentSlideRels.serialize(),
      path: `ppt/slides/_rels/slide${i + 1}.xml.rels`,
    };
  }

  // Content Types — charts + smartarts
  const chartCountFromXmlComponents = charts.array.length;
  for (let i = 0; i < chartCountFromXmlComponents; i++) contentTypes.addChart(i + 1);
  for (let i = 0; i < descCtx.charts.length; i++)
    contentTypes.addChart(chartCountFromXmlComponents + i + 1);

  const saCountFromXmlComponents = smartArts.array.length;
  for (let i = 0; i < saCountFromXmlComponents; i++) {
    contentTypes.addDiagramData(i + 1);
    contentTypes.addDiagramLayout(i + 1);
    contentTypes.addDiagramStyle(i + 1);
    contentTypes.addDiagramColors(i + 1);
    contentTypes.addDiagramDrawing(i + 1);
  }
  for (let i = 0; i < descCtx.smartArts.length; i++) {
    const idx = saCountFromXmlComponents + i + 1;
    contentTypes.addDiagramData(idx);
    contentTypes.addDiagramLayout(idx);
    contentTypes.addDiagramStyle(idx);
    contentTypes.addDiagramColors(idx);
    contentTypes.addDiagramDrawing(idx);
  }
  mapping["ContentTypes"] = {
    data: XML_DECL + (contentTypesDesc.stringify({ builder: contentTypes }, descCtx) ?? ""),
    path: "[Content_Types].xml",
  };

  // Compile mapping to Zippable
  const files = compileMapping(mapping, overrides);

  // Chart parts
  const allCharts = [
    ...charts.array.map((c) => ({
      key: c.key,
      xml: XML_DECL + c.chartSpaceXml,
    })),
    ...descCtx.charts.map((c) => ({ key: c.key, xml: c.chartSpaceXml })),
  ];
  for (let i = 0; i < allCharts.length; i++) {
    files[`ppt/charts/chart${i + 1}.xml`] = encoder.encode(allCharts[i].xml);
    files[`ppt/charts/_rels/chart${i + 1}.xml.rels`] = encoder.encode(
      `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
    );
  }

  // SmartArt parts
  const allSmartArts = [
    ...smartArts.array.map((s) => ({
      key: s.key,
      dataModelXml: XML_DECL + s.dataModelXml,
      layout: s.layout,
      style: s.style,
      color: s.color,
    })),
    ...descCtx.smartArts.map((s) => ({
      key: s.key,
      dataModelXml: s.dataModelXml,
      layout: s.layout,
      style: s.style,
      color: s.color,
    })),
  ];
  for (let i = 0; i < allSmartArts.length; i++) {
    const sa = allSmartArts[i];
    files[`ppt/diagrams/data${i + 1}.xml`] = encoder.encode(sa.dataModelXml);
    files[`ppt/diagrams/layout${i + 1}.xml`] = encoder.encode(getLayoutXml(sa.layout));
    files[`ppt/diagrams/quickStyle${i + 1}.xml`] = encoder.encode(getStyleXml(sa.style));
    files[`ppt/diagrams/colors${i + 1}.xml`] = encoder.encode(getColorXml(sa.color));
    files[`ppt/diagrams/drawing${i + 1}.xml`] = encoder.encode(DEFAULT_DRAWING_XML);
  }

  // ViewProps relationships
  if (hasOutlineViewSlides) {
    const vpRels = new Relationships();
    for (let i = 0; i < slides.length; i++) {
      vpRels.addRelationship(
        i + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        `slides/slide${i + 1}.xml`,
      );
    }
    files["ppt/_rels/viewProps.xml.rels"] = encoder.encode(XML_DECL + vpRels.serialize());
  }

  // PresProps relationships
  if (htmlPublishInfo) {
    const presPropsRels = new Relationships();
    presPropsRels.addRelationship(
      htmlPublishInfo.rId.replace("rId", ""),
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      htmlPublishInfo.target ?? "presentation.htm",
      "External",
    );
    files["ppt/_rels/presProps.xml.rels"] = encoder.encode(XML_DECL + presPropsRels.serialize());
  }

  // Notes slides
  const notesSlideToSlide = new Map<number, number>();
  for (const [slideIdx, notesIdx] of notesSlideIndexMap) {
    notesSlideToSlide.set(notesIdx, slideIdx);
  }
  for (let i = 0; i < notesTexts.length; i++) {
    files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = encoder.encode(
      XML_DECL + (notesSlideDesc.stringify({ text: notesTexts[i] }, descCtx) ?? ""),
    );
    const slideIdx = notesSlideToSlide.get(i) ?? 0;
    const nsRels = new Relationships();
    nsRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
      "../notesMasters/notesMaster1.xml",
    );
    nsRels.addRelationship(
      2,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
      `../slides/slide${slideIdx + 1}.xml`,
    );
    files[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = encoder.encode(
      XML_DECL + nsRels.serialize(),
    );
  }

  // Slide sync properties
  for (let i = 0; i < slideSyncOptionsList.length; i++) {
    files[`ppt/slideSyncPr/slideSyncPr${i + 1}.xml`] = encoder.encode(
      XML_DECL + (slideSyncDesc.stringify(slideSyncOptionsList[i], descCtx) ?? ""),
    );
  }

  // Comment authors
  if (commentAuthorEntries) {
    files["ppt/commentAuthors.xml"] = encoder.encode(
      XML_DECL + (commentAuthorsDesc.stringify(commentAuthorEntries, descCtx) ?? ""),
    );
  }

  // Slide comments
  for (let i = 0; i < slideCommentEntries.length; i++) {
    if (slideCommentEntries[i]) {
      files[`ppt/comments/comment${i + 1}.xml`] = encoder.encode(
        XML_DECL + (slideCommentsDesc.stringify(slideCommentEntries[i]!, descCtx) ?? ""),
      );
    }
  }

  // Media files
  for (const image of media.array) {
    files[`ppt/media/${image.fileName}`] = [
      image.data,
      { level: levelForMediaName(image.fileName, mediaLevel) as ZipOptions["level"] },
    ];
    if (image.type === "svg" && "fallback" in image) {
      const fallback = (
        image as MediaData & {
          fallback: { fileName: string; data: Uint8Array };
        }
      ).fallback;
      files[`ppt/media/${fallback.fileName}`] = [
        fallback.data,
        { level: levelForMediaName(fallback.fileName, mediaLevel) as ZipOptions["level"] },
      ];
    }
  }

  return files;
}

function getChartGlobalIndex(
  key: string,
  legacyCharts: { key: string }[],
  descCharts: { key: string }[],
): number {
  const legacyIdx = legacyCharts.findIndex((c) => c.key === key);
  if (legacyIdx >= 0) return legacyIdx;
  return legacyCharts.length + descCharts.findIndex((c) => c.key === key);
}

function computeSmartArtGlobalStart(
  firstKey: string,
  legacySmartArts: { key: string }[],
  descSmartArts: { key: string }[],
): number {
  const legacyIdx = legacySmartArts.findIndex((s) => s.key === firstKey);
  if (legacyIdx >= 0) return legacyIdx;
  return legacySmartArts.length + descSmartArts.findIndex((s) => s.key === firstKey);
}
