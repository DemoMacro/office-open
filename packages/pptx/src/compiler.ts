/**
 * Descriptor-based PPTX compiler.
 *
 * Produces a valid PPTX ZIP archive using the descriptor pipeline.
 * Accepts pure JSON PresentationOptions — no intermediate class needed.
 *
 * @module
 */

import {
  IMAGE_MEDIA_CONTENT_TYPES,
  Relationships,
  addBinaryFile,
  buildRootRelationships,
  convertToEmu,
} from "@office-open/core";
import type { RelationshipType } from "@office-open/core";
import {
  appPropertiesDesc,
  buildCorePropertiesXmlString,
  collectPlaceholderKeys,
  compileMapping,
  customPropertiesDesc,
  getReferencedMedia,
  getAudioRefs,
  getMediaRefs,
  getOleRefs,
  getVideoRefs,
  hasPlaceholders,
  replaceAudioPlaceholders,
  replaceChartPlaceholders,
  replaceHyperlinkPlaceholders,
  replaceImagePlaceholders,
  replaceMediaPlaceholders,
  replaceOlePlaceholders,
  replaceSmartArtPlaceholders,
  replaceVideoPlaceholders,
  addSmartArtRelationships,
  contentTypesDesc,
  deriveContentTypes,
  resolverFromRegistry,
  themeOverrideDesc,
  toUint8Array,
  PPTX_PARTS,
} from "@office-open/core";
import type { XmlifyedFile, Zippable } from "@office-open/core";
import { ChartCollection } from "@office-open/core/chart";
import { SmartArtCollection } from "@office-open/core/smartart";
import {
  stringifyColorDefinitionPart,
  stringifyLayoutDefinitionPart,
  stringifyStyleDefinitionPart,
} from "@office-open/core/smartart";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import type { AuthorEntry, CommentEntry } from "@parts/comment";
import type { PresentationPartOptions, PresentationSectionGroup } from "@parts/presentation";
import { buildCustomLayoutXml, buildLayoutXml, type SlideLayoutType } from "@parts/slide-layout";
import { stringifyControls, stringifyCustDataLst } from "@parts/slide/c-sld";
import type { SlideSyncOptions } from "@parts/slide/slide-sync-properties";
import { getColorXml, getLayoutXml, getStyleXml, DEFAULT_DRAWING_XML } from "@parts/smartart";
import { SP_TREE_HEADER } from "@shared/constants";
import {
  type PresentationOptions,
  type MasterDefinition,
  type LayoutDefinition,
  type SlideOptions,
  type SlideSize,
} from "@shared/file";
import { buildHeaderFooterShapes } from "@shared/header-footer";
import type { MediaData } from "@shared/media/data";
import { createThemeXml } from "@shared/theme";
import { buildTransition } from "@shared/transition";

import { PptxWriteContext } from "./context";
import { timingDesc } from "./parts/descriptors/animation";
import { backgroundDesc } from "./parts/descriptors/background";
import { stringifyChild } from "./parts/descriptors/bridge";
import { colorMappingOverrideDesc } from "./parts/descriptors/color-map-override";
import { commentAuthorsDesc, slideCommentsDesc } from "./parts/descriptors/comments";
import { handoutMasterDesc } from "./parts/descriptors/handout-master";
import { notesMasterDesc } from "./parts/descriptors/notes-master";
import { notesSlideDesc, type NotesSlideOptions } from "./parts/descriptors/notes-slide";
import { presentationDesc } from "./parts/descriptors/presentation";
import { presentationPropertiesDesc } from "./parts/descriptors/presentation-properties";
import { parseLayoutDef, slideLayoutDesc } from "./parts/descriptors/slide-layout";
import { slideMasterDesc } from "./parts/descriptors/slide-master";
import { slideSyncDesc } from "./parts/descriptors/slide-sync";
import { tableStylesDesc } from "./parts/descriptors/table-styles";
import { viewPropsDesc } from "./parts/descriptors/view-properties";

// ── Constants ──

const encoder = new TextEncoder();
const XML_DECL = OOXML_XML_DECLARATION + "\n";

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
  def: LayoutDefinition;
  /** Serialized themeOverride part XML, when the layout deviates from its master's theme. */
  themeOverride?: string;
}

interface MasterInfo {
  name: string;
  index: number;
  master: string;
  theme: string;
  /** Emitted theme part index — shared when masters carry identical themes. */
  themeIndex: number;
  layouts: LayoutInfo[];
  masterRels: Relationships;
  layoutRels: Relationships[];
}

/**
 * Resolve a layout definition to its structured form for stringify.
 *
 * Structured defs (round-trip parse, or user-provided shapes/bg/transition/etc)
 * pass through unchanged. Fresh template / custom / deprecated-verbatim layouts
 * are built to XML then parsed back to structure, so every layout is emitted via
 * slideLayoutDesc.stringify uniformly.
 */
function resolveLayoutDef(
  layoutDef: LayoutDefinition | undefined,
  slideLayoutType: SlideLayoutType,
  slideWidth: number,
): LayoutDefinition {
  if (layoutDef && hasStructuredLayoutContent(layoutDef)) return layoutDef;
  const xml = layoutDef?.layout
    ? layoutDef.layout
    : layoutDef
      ? buildCustomLayoutXml(layoutDef)
      : buildLayoutXml(slideLayoutType, slideWidth);
  return parseLayoutDef(xml);
}

/** True when a def carries structured content that must drive stringify directly. */
function hasStructuredLayoutContent(def: LayoutDefinition): boolean {
  return (
    (def.children !== undefined && def.children.length > 0) ||
    def.background !== undefined ||
    def.transition !== undefined ||
    def.animations !== undefined ||
    def.ext !== undefined ||
    def.headerFooter !== undefined ||
    def.colorMappingOverride !== undefined ||
    (def.controls !== undefined && def.controls.length > 0) ||
    (def.customerData !== undefined && def.customerData.length > 0)
  );
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
  return { width: convertToEmu(size.width), height: convertToEmu(size.height) };
}

function deriveInitials(name: string): string {
  const parts = name.trim().split(/\s+/);
  if (parts.length < 2) return name.slice(0, 2).toUpperCase();
  const first = parts[0];
  const last = parts[parts.length - 1];
  if (!first || !last) return name.slice(0, 2).toUpperCase();
  return (first.charAt(0) + last.charAt(0)).toUpperCase();
}

function buildMasterMap(
  masterDefs: MasterDefinition[],
  slides: SlideOptions[],
  slideWidth: number,
  ctx: PptxWriteContext,
  passthroughRelationships: PresentationOptions["passthroughRelationships"],
): MasterInfo[] {
  // Master placeholder positions scale to the slide width — record it on the
  // shared context so slideMasterDesc.stringify can read it.
  ctx.slideWidth = slideWidth;
  const defs = masterDefs.length > 0 ? masterDefs : [{} as MasterDefinition];
  const slideMasterLookup = new Map<number, number>();

  for (const [si, slide] of slides.entries()) {
    const masterName = slide.master;
    if (masterName === undefined) {
      slideMasterLookup.set(si, 0);
      continue;
    }
    const mi = defs.findIndex((d) => d.name === masterName);
    slideMasterLookup.set(si, mi >= 0 ? mi : 0);
  }

  let globalLayoutIndex = 0;
  const masters: MasterInfo[] = [];
  // Identical master themes share one theme part (sources commonly point two
  // masters at the same theme) — dedupe by serialized content, like media.
  const themeIndexByXml = new Map<string, number>();
  let themeCount = 0;

  for (const [mi, def] of defs.entries()) {
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
      for (const [si, slide] of slides.entries()) {
        if (slideMasterLookup.get(si) === mi) {
          const lt = slide.layout ?? "blank";
          if (!seen.has(lt)) {
            seen.add(lt);
            keys.push(lt);
          }
        }
      }
      layoutKeys = keys.length > 0 ? keys : ["blank"];
    }

    // Layout id + rId pairs — rId order matches masterRels below.
    const layoutIdBase = 2147483648 + mi * 12 + 1;
    const slideLayoutIds = layoutKeys.map((_, li) => ({
      id: layoutIdBase + li,
      relationshipId: `rId${li + 1}`,
    }));
    // Rest-spread so every SlideMasterOptions field flows to the descriptor —
    // field-copy whitelists here have dropped newly added options before.
    const { name: _masterName, theme: _theme, layouts: _layouts, ...masterOpts } = def;
    const master = slideMasterDesc.stringify({ ...masterOpts, slideLayoutIds }, ctx) ?? "";
    const theme = createThemeXml(def.theme, ctx);
    let themeIndex = themeIndexByXml.get(theme);
    if (themeIndex === undefined) {
      themeIndex = themeCount++;
      themeIndexByXml.set(theme, themeIndex);
    }

    const layouts: LayoutInfo[] = [];
    const layoutRels: Relationships[] = [];

    for (const [li, key] of layoutKeys.entries()) {
      const layoutDef = layoutDefs?.[li];
      const slideLayoutType = (layoutDef?.type ?? key) as SlideLayoutType;
      const themeOverride = layoutDef?.themeOverride
        ? (themeOverrideDesc.stringify(layoutDef.themeOverride, ctx) ?? undefined)
        : undefined;
      layouts.push({
        key,
        index: globalLayoutIndex,
        masterIndex: mi,
        def: resolveLayoutDef(layoutDef, slideLayoutType, slideWidth),
        themeOverride,
      });
      const layoutRelEntries: RelEntry[] = [
        {
          id: 1,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
          target: `../slideMasters/slideMaster${mi + 1}.xml`,
        },
      ];
      if (themeOverride) {
        layoutRelEntries.push({
          id: 2,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride",
          target: `../theme/themeOverride${globalLayoutIndex + 1}.xml`,
        });
      }
      const layoutRel = buildRels(layoutRelEntries);
      // Layout-level passthrough relationships (round-trip) — re-emitted as
      // written unless the model already registered the same kind (ownership
      // test: targets may be renamed and ISO-strict types differ in URI only).
      for (const rel of passthroughRelationships ?? []) {
        if (rel.source !== `ppt/slideLayouts/slideLayout${globalLayoutIndex + 1}.xml`) continue;
        if (layoutRel.hasRelationshipKind(rel.relationshipType.split("/").pop()!)) continue;
        if (layoutRel.hasRelationship(rel.relationshipType, rel.target)) continue;
        layoutRel.add(rel.relationshipType as RelationshipType, rel.target);
      }
      layoutRels.push(layoutRel);
      globalLayoutIndex++;
    }

    const masterRelsEntries: RelEntry[] = [];
    for (const [li, layout] of layouts.entries()) {
      masterRelsEntries.push({
        id: li + 1,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        target: `../slideLayouts/slideLayout${layout.index + 1}.xml`,
      });
    }
    masterRelsEntries.push({
      id: layouts.length + 1,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      target: `../theme/theme${themeIndex + 1}.xml`,
    });
    // Master-level passthrough relationships (round-trip) — re-emitted as
    // written unless the model already registered the same kind (ownership
    // test: targets may be renamed and ISO-strict types differ in URI only).
    for (const rel of passthroughRelationships ?? []) {
      if (rel.source !== `ppt/slideMasters/slideMaster${mi + 1}.xml`) continue;
      const kind = rel.relationshipType.split("/").pop();
      const kindOwned = masterRelsEntries.some((e) => e.type.split("/").pop() === kind);
      const exists = masterRelsEntries.some(
        (e) => e.type.split("/").pop() === kind && e.target === rel.target,
      );
      if (kindOwned || exists) continue;
      masterRelsEntries.push({
        id: masterRelsEntries.length + 1,
        type: rel.relationshipType as RelationshipType,
        target: rel.target,
      });
    }

    masters.push({
      name,
      index: mi,
      master,
      theme,
      themeIndex,
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
  // slideIndex is caller-bounded by slides.length; masters is built by buildMasterMap
  // with non-empty layouts — these indexed accesses are contract narrows, not runtime checks.
  const opts = slides[slideIndex]!;
  const mi =
    opts.master !== undefined
      ? Math.max(
          0,
          masters.findIndex((m) => m.name === opts.master),
        )
      : 0;
  const master = masters[mi]!;
  const layoutKey = opts.layout ?? "blank";
  const li = master.layouts.find((l) => l.key === layoutKey);
  return li ?? master.layouts[0]!;
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

export function buildCommentData(
  slides: SlideOptions[],
  existingAuthors: AuthorEntry[] = [],
): {
  authors: AuthorEntry[] | undefined;
  perSlide: (CommentEntry[] | undefined)[];
} {
  const authorMap = new Map<
    string,
    { id: number; name: string; initials: string; clrIdx: number; commentCount: number }
  >();
  let nextAuthorId = 0;
  // Seed from existing authors so appended comments continue author ids and the
  // per-author idx counter (commentCount resumes at lastIdx).
  for (const a of existingAuthors) {
    authorMap.set(a.name, {
      id: a.id,
      name: a.name,
      initials: a.initials,
      clrIdx: a.clrIdx,
      commentCount: a.lastIdx,
    });
    if (a.id >= nextAuthorId) nextAuthorId = a.id + 1;
  }

  const perSlide: (CommentEntry[] | undefined)[] = Array.from({ length: slides.length });

  for (const [i, slide] of slides.entries()) {
    const slideComments = slide.comments;
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
        x: convertToEmu(c.x),
        y: convertToEmu(c.y),
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

/** PPTX part path → content type, derived from the part registry. Matches
 * actual file paths, so dense (slides) and sparse (slide-indexed comments)
 * naming are both handled. */
const PPTX_CONTENT_TYPE_RESOLVER = resolverFromRegistry(PPTX_PARTS);

/** Extension → MIME for media Default entries (image/video/audio). Declared
 * only for extensions actually present in the package. */
const PPTX_MEDIA_CONTENT_TYPES: Record<string, string> = {
  ...IMAGE_MEDIA_CONTENT_TYPES,
  mp4: "video/mp4",
  mov: "video/quicktime",
  wmv: "video/x-ms-wmv",
  avi: "video/x-msvideo",
  mpg: "video/mpeg",
  mpeg: "video/mpeg",
  mp3: "audio/mpeg",
  wav: "audio/wav",
  wma: "audio/x-ms-wma",
  aac: "audio/aac",
  bin: "application/vnd.openxmlformats-officedocument.oleObject",
};

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
  // presentation.xml.rels carries exactly one theme rel (the presentation's
  // default theme) regardless of how many masters or theme parts exist —
  // extra masters reference their themes through their own rels.
  rels.addRelationship(
    rid++,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "theme/theme1.xml",
  );
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
    | "smartTags"
    | "defaultTextStyle"
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
    options.customerData === undefined &&
    options.smartTags === undefined &&
    options.defaultTextStyle === undefined
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
    smartTags: options.smartTags,
    defaultTextStyle: options.defaultTextStyle,
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
  if (slideOpts.hidden) sldAttrs.push(' show="0"');
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

  // Per-slide header/footer: instantiate the dt/ftr/sldNum placeholder shapes
  // after the children (spTree tail, ids continue the child sequence). A type
  // the children already carry — a round-tripped placeholder shape — is left
  // untouched so re-serialization never duplicates it.
  if (slideOpts.headerFooter) {
    const present = new Set(
      (slideOpts.children ?? []).flatMap((c) =>
        "shape" in c && c.shape?.placeholder ? [c.shape.placeholder] : [],
      ),
    );
    for (const shape of buildHeaderFooterShapes(slideOpts.headerFooter)) {
      if (present.has(shape.placeholder!)) continue;
      const xml = stringifyChild({ shape }, ctx);
      if (xml) parts.push(xml);
    }
  }

  parts.push("</p:spTree>");

  parts.push(stringifyCustDataLst(slideOpts.customerData));
  parts.push(stringifyControls(slideOpts.controls));

  // cSld-tail extLst (p14:creationId's home) — verbatim, before the cSld close
  // and distinct from the root-level extLst emitted after timing.
  if (slideOpts.cSldExt) {
    parts.push(`<p:extLst>${slideOpts.cSldExt}</p:extLst>`);
  }

  parts.push("</p:cSld>");
  parts.push(
    colorMappingOverrideDesc.stringify(slideOpts.colorMappingOverride, ctx) ??
      "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>",
  );

  if (slideOpts.transition) {
    parts.push(buildTransition(slideOpts.transition, ctx));
  }

  if (slideOpts.animations && slideOpts.animations.length > 0) {
    parts.push(timingDesc.stringify(slideOpts.animations, ctx) ?? "");
  }

  // p:extLst — verbatim round-trip (last child per CT_Slide sequence)
  if (slideOpts.ext) {
    parts.push(`<p:extLst>${slideOpts.ext}</p:extLst>`);
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

  const masters = buildMasterMap(
    masterDefs,
    slides,
    sz.width,
    descCtx,
    options.passthroughRelationships,
  );
  const allLayouts = masters.flatMap((m) => m.layouts);
  const allLayoutRels = masters.flatMap((m) => m.layoutRels);
  // Unique master themes in theme-index order (deduped in buildMasterMap) —
  // notesMaster/handoutMaster themes append after these.
  const uniqueMasterThemes: string[] = [];
  for (const m of masters) uniqueMasterThemes[m.themeIndex] ??= m.theme;
  const themes = uniqueMasterThemes;
  const masterRels = masters.map((m) => m.masterRels);
  const slideRels = buildSlideRels(masters, slides);
  const { authors: commentAuthorEntries, perSlide: slideCommentEntries } = buildCommentData(slides);

  const notesOptions: NotesSlideOptions[] = [];
  const notesSlideIndexMap = new Map<number, number>();
  let notesIdx = 0;
  for (const [i, slide] of slides.entries()) {
    if (slide.notes) {
      notesOptions.push(typeof slide.notes === "string" ? { text: slide.notes } : slide.notes);
      notesSlideIndexMap.set(i, notesIdx++);
    }
  }

  const slideSyncOptionsList: SlideSyncOptions[] = [];
  const slideSyncIndexMap = new Map<number, number>();
  let syncIdx = 0;
  for (const [i, slide] of slides.entries()) {
    if (slide.slideSync) {
      slideSyncOptionsList.push(slide.slideSync);
      slideSyncIndexMap.set(i, syncIdx++);
    }
  }

  // ── Mutable state ──

  const hasCustomProperties = !!options.customProperties && options.customProperties.length > 0;
  const presRels = initPresRels(masters, slides.length);
  // Group slides into p14:sections by name (first-occurrence order); slides
  // without a section name are left ungrouped (absent from p14:sectionLst).
  const sectionOrder: string[] = [];
  const sectionIndices = new Map<string, number[]>();
  for (const [i, slide] of slides.entries()) {
    const name = slide.section;
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
    ...(options.ext !== undefined ? { ext: options.ext } : {}),
  };
  const fileRels = buildRootRelationships(
    "ppt/presentation.xml",
    hasCustomProperties,
    options.passthroughRelationships,
  );
  const media = descCtx.mediaCollection;
  const charts = new ChartCollection();
  const smartArts = new SmartArtCollection();

  const presPropsFullOpts =
    options.web ||
    options.print ||
    options.htmlPublish ||
    options.colorMru ||
    options.show ||
    options.presentationPropertiesExt
      ? {
          web: options.web,
          print: options.print,
          htmlPublish: options.htmlPublish,
          colorMru: options.colorMru,
          show: options.show,
          ext: options.presentationPropertiesExt,
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
    data: XML_DECL + (presentationPropertiesDesc.stringify(presPropsFullOpts ?? {}, descCtx) ?? ""),
    path: "ppt/presProps.xml",
  };

  mapping["ViewProps"] = {
    data: XML_DECL + (viewPropsDesc.stringify(options.view ?? {}, descCtx) ?? ""),
    path: "ppt/viewProps.xml",
  };

  // Slide Masters
  for (const [mi, masterInfo] of masters.entries()) {
    mapping[`SlideMaster${mi}`] = {
      data: XML_DECL + masterInfo.master,
      path: `ppt/slideMasters/slideMaster${mi + 1}.xml`,
    };
    mapping[`SlideMasterRels${mi}`] = {
      data: XML_DECL + masterRels[mi]!.serialize(),
      path: `ppt/slideMasters/_rels/slideMaster${mi + 1}.xml.rels`,
    };
  }

  // Slide Layouts
  for (const [li, layoutInfo] of allLayouts.entries()) {
    const layoutXml = slideLayoutDesc.stringify(layoutInfo.def, descCtx) ?? "";
    // Media referenced by layout shapes gets the same image-relationship
    // wiring slides use (layout pictures otherwise lose their rel).
    const layoutRels = allLayoutRels[li]!;
    const layoutMediaData = getReferencedMedia(layoutXml, media.array);
    const layoutImageOffset = layoutRels.relationshipCount + 1;
    for (const [idx, mediaItem] of layoutMediaData.entries()) {
      layoutRels.addRelationship(
        layoutImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${mediaItem.fileName}`,
      );
    }
    mapping[`SlideLayout${li}`] = {
      data: XML_DECL + replaceImagePlaceholders(layoutXml, layoutMediaData, layoutImageOffset),
      path: `ppt/slideLayouts/slideLayout${li + 1}.xml`,
    };
    mapping[`SlideLayoutRels${li}`] = {
      data: XML_DECL + layoutRels.serialize(),
      path: `ppt/slideLayouts/_rels/slideLayout${li + 1}.xml.rels`,
    };
    if (layoutInfo.themeOverride) {
      mapping[`SlideLayoutThemeOverride${li}`] = {
        data: XML_DECL + layoutInfo.themeOverride,
        path: `ppt/theme/themeOverride${li + 1}.xml`,
      };
    }
  }

  // Notes Master — emitted when notes slides exist or the source carried one.
  const includeNotesMasterPart =
    notesOptions.length > 0 || options.notesMasterOptions !== undefined;
  if (includeNotesMasterPart) {
    const notesMasterRId = presRels.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
      "notesMasters/notesMaster1.xml",
    );
    presOptions.notesMasterRId = notesMasterRId;
    const notesMasterThemeIndex = themes.length + 1;
    const notesMasterXml =
      notesMasterDesc.stringify(options.notesMasterOptions ?? {}, descCtx) ?? "";
    const notesMasterThemeXml = createThemeXml(options.notesMasterOptions?.theme, descCtx);
    mapping["NotesMasterTheme"] = {
      data: XML_DECL + notesMasterThemeXml,
      path: `ppt/theme/theme${notesMasterThemeIndex}.xml`,
    };
    const notesMasterRels = new Relationships();
    notesMasterRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      `../theme/theme${notesMasterThemeIndex}.xml`,
    );
    // Media referenced by notes-master shapes gets slide-style image wiring.
    const notesMasterMediaData = getReferencedMedia(notesMasterXml, media.array);
    const notesMasterImageOffset = notesMasterRels.relationshipCount + 1;
    for (const [idx, mediaItem] of notesMasterMediaData.entries()) {
      notesMasterRels.addRelationship(
        notesMasterImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${mediaItem.fileName}`,
      );
    }
    mapping["NotesMaster"] = {
      data:
        XML_DECL +
        replaceImagePlaceholders(notesMasterXml, notesMasterMediaData, notesMasterImageOffset),
      path: "ppt/notesMasters/notesMaster1.xml",
    };
    mapping["NotesMasterRelationships"] = {
      data: XML_DECL + notesMasterRels.serialize(),
      path: "ppt/notesMasters/_rels/notesMaster1.xml.rels",
    };
  }

  // Handout Master
  if (includeHandout) {
    const handoutMasterRId = presRels.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster",
      "handoutMasters/handoutMaster1.xml",
    );
    presOptions.handoutMasterRId = handoutMasterRId;
    const handoutMasterThemeIndex = themes.length + (includeNotesMasterPart ? 2 : 1);
    const handoutMasterXml =
      handoutMasterDesc.stringify({ options: options.handoutMasterOptions }, descCtx) ?? "";
    const handoutMasterThemeXml = createThemeXml(options.handoutMasterOptions?.theme, descCtx);
    mapping["HandoutMasterTheme"] = {
      data: XML_DECL + handoutMasterThemeXml,
      path: `ppt/theme/theme${handoutMasterThemeIndex}.xml`,
    };
    const handoutMasterRels = new Relationships();
    handoutMasterRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      `../theme/theme${handoutMasterThemeIndex}.xml`,
    );
    // Media referenced by handout-master shapes gets slide-style image wiring.
    const handoutMediaData = getReferencedMedia(handoutMasterXml, media.array);
    const handoutImageOffset = handoutMasterRels.relationshipCount + 1;
    for (const [idx, mediaItem] of handoutMediaData.entries()) {
      handoutMasterRels.addRelationship(
        handoutImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${mediaItem.fileName}`,
      );
    }
    mapping["HandoutMaster"] = {
      data:
        XML_DECL + replaceImagePlaceholders(handoutMasterXml, handoutMediaData, handoutImageOffset),
      path: "ppt/handoutMasters/handoutMaster1.xml",
    };
    mapping["HandoutMasterRelationships"] = {
      data: XML_DECL + handoutMasterRels.serialize(),
      path: "ppt/handoutMasters/_rels/handoutMaster1.xml.rels",
    };
  }

  // Comment Authors
  if (commentAuthorEntries) {
    presRels.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
      "commentAuthors.xml",
    );
  }

  // Presentation XML
  const presBody = presentationDesc.stringify(presOptions, descCtx);
  const presentationXml = presBody ? XML_DECL + presBody : "";
  const mediaData = getReferencedMedia(presentationXml, media.array);
  const presImageOffset = presRels.relationshipCount + 1;
  for (const [idx, mediaItem] of mediaData.entries()) {
    presRels.addRelationship(
      presImageOffset + idx,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      `../media/${mediaItem.fileName}`,
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
  // Passthrough relationships (round-trip): the source presentation.xml.rels
  // referenced parts the model carries verbatim (handoutMaster, customXml, …).
  // Re-emit them as written — targets are passthrough paths that never move.
  for (const rel of options.passthroughRelationships ?? []) {
    if (rel.source !== "ppt/presentation.xml") continue;
    if (presRels.hasRelationship(rel.relationshipType, rel.target)) continue;
    presRels.add(rel.relationshipType as RelationshipType, rel.target);
  }
  mapping["PresentationRelationships"] = {
    data: XML_DECL + presRels.serialize(),
    path: "ppt/_rels/presentation.xml.rels",
  };

  // Slides
  for (const [i, slide] of slides.entries()) {
    const slideXml = stringifySlide(slide, descCtx);

    const slideMediaData = getReferencedMedia(slideXml, media.array);
    const currentSlideRels = slideRels[i];
    if (!currentSlideRels) continue; // slideRels is built one-per-slide in lockstep with slides
    const slideImageOffset = currentSlideRels.relationshipCount + 1;
    for (const [idx, mediaItem] of slideMediaData.entries()) {
      currentSlideRels.addRelationship(
        slideImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${mediaItem.fileName}`,
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
        for (const [ci, chartKey] of allChartKeys.entries()) {
          currentSlideRels.addRelationship(
            slideChartOffset + ci,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            `../charts/chart${getChartGlobalIndex(chartKey, charts.array, descCtx.charts) + 1}.xml`,
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
        const firstSaKey = allSaKeys[0];
        if (firstSaKey !== undefined) {
          const saGlobalStart = computeSmartArtGlobalStart(
            firstSaKey,
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
      }

      // Hyperlinks
      const slideHlinkKeys = collectPlaceholderKeys(replacedSlideXml, "hlink:");
      if (slideHlinkKeys.length > 0) {
        const slideHlinkKeySet = new Set(slideHlinkKeys);
        const slideHlinks = descCtx.hyperlinks.filter((h) => slideHlinkKeySet.has(h.key));
        const hlinkOffset = currentSlideRels.relationshipCount + 1;
        replacedSlideXml = replaceHyperlinkPlaceholders(replacedSlideXml, slideHlinks, hlinkOffset);
        for (const [hi, hlink] of slideHlinks.entries()) {
          if (hlink.slide !== undefined) {
            // Internal slide jump: r:id → slideN.xml, no TargetMode (internal).
            currentSlideRels.addRelationship(
              hlinkOffset + hi,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
              `slide${hlink.slide}.xml`,
            );
          } else {
            currentSlideRels.addRelationship(
              hlinkOffset + hi,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
              hlink.url ?? "",
              "External",
            );
          }
        }
      }

      // Media (video/audio)
      const slideMediaRefs = getMediaRefs(replacedSlideXml, media.array);
      const slideAudioRefs = getAudioRefs(replacedSlideXml, media.array);
      const slideVideoRefs = getVideoRefs(replacedSlideXml, media.array);
      if (slideMediaRefs.length > 0 || slideAudioRefs.length > 0 || slideVideoRefs.length > 0) {
        const mediaOffset = currentSlideRels.relationshipCount + 1;
        const audioOffset = mediaOffset + slideMediaRefs.length;
        const videoOffset = audioOffset + slideAudioRefs.length;
        replacedSlideXml = replaceMediaPlaceholders(replacedSlideXml, slideMediaRefs, mediaOffset);
        replacedSlideXml = replaceAudioPlaceholders(replacedSlideXml, slideAudioRefs, audioOffset);
        replacedSlideXml = replaceVideoPlaceholders(replacedSlideXml, slideVideoRefs, videoOffset);
        for (const [mi, mediaRef] of slideMediaRefs.entries()) {
          currentSlideRels.addRelationship(
            mediaOffset + mi,
            "http://schemas.microsoft.com/office/2007/relationships/media",
            `../media/${mediaRef.fileName}`,
          );
        }
        for (const [ai, audioRef] of slideAudioRefs.entries()) {
          currentSlideRels.addRelationship(
            audioOffset + ai,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
            `../media/${audioRef.fileName}`,
          );
        }
        for (const [vi, videoRef] of slideVideoRefs.entries()) {
          currentSlideRels.addRelationship(
            videoOffset + vi,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
            `../media/${videoRef.fileName}`,
          );
        }
      }

      // OLE embeddings
      const slideOleRefs = getOleRefs(replacedSlideXml, descCtx.embeddings);
      if (slideOleRefs.length > 0) {
        const oleOffset = currentSlideRels.relationshipCount + 1;
        replacedSlideXml = replaceOlePlaceholders(replacedSlideXml, slideOleRefs, oleOffset);
        for (const [oi, oleRef] of slideOleRefs.entries()) {
          currentSlideRels.addRelationship(
            oleOffset + oi,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
            `../embeddings/${oleRef.fileName}`,
          );
        }
      }
    }

    mapping[`Slide${i}`] = {
      data: replacedSlideXml,
      path: `ppt/slides/slide${i + 1}.xml`,
    };

    if (slideCommentEntries[i]) {
      currentSlideRels.add(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        `../comments/comment${i + 1}.xml`,
      );
    }

    const notesSlideIndex = notesSlideIndexMap.get(i);
    if (notesSlideIndex !== undefined) {
      currentSlideRels.add(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
        `../notesSlides/notesSlide${notesSlideIndex + 1}.xml`,
      );
    }

    const slideSyncIndex = slideSyncIndexMap.get(i);
    if (slideSyncIndex !== undefined) {
      currentSlideRels.add(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideSyncProperties",
        `../slideSyncPr/slideSyncPr${slideSyncIndex + 1}.xml`,
      );
    }

    // Slide-level passthrough relationships (round-trip), appended last so the
    // kind ownership test sees every model registration above: a source rel
    // whose kind the model already registered (comments, notesSlide, image
    // under a renamed file, …) is skipped; the rest (tags, VML drawings, OLE
    // embeddings the model did not absorb, …) re-emit as written.
    for (const rel of options.passthroughRelationships ?? []) {
      if (rel.source !== `ppt/slides/slide${i + 1}.xml`) continue;
      if (currentSlideRels.hasRelationshipKind(rel.relationshipType.split("/").pop()!)) continue;
      if (currentSlideRels.hasRelationship(rel.relationshipType, rel.target)) continue;
      currentSlideRels.add(rel.relationshipType as RelationshipType, rel.target);
    }

    mapping[`SlideRelationships${i}`] = {
      data: XML_DECL + currentSlideRels.serialize(),
      path: `ppt/slides/_rels/slide${i + 1}.xml.rels`,
    };
  }

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
  for (const [i, chart] of allCharts.entries()) {
    files[`ppt/charts/chart${i + 1}.xml`] = encoder.encode(chart.xml);
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
  for (const [i, sa] of allSmartArts.entries()) {
    files[`ppt/diagrams/data${i + 1}.xml`] = encoder.encode(sa.dataModelXml);
    files[`ppt/diagrams/layout${i + 1}.xml`] = encoder.encode(
      typeof sa.layout === "string"
        ? getLayoutXml(sa.layout)
        : stringifyLayoutDefinitionPart(sa.layout),
    );
    files[`ppt/diagrams/quickStyle${i + 1}.xml`] = encoder.encode(
      typeof sa.style === "string" ? getStyleXml(sa.style) : stringifyStyleDefinitionPart(sa.style),
    );
    files[`ppt/diagrams/colors${i + 1}.xml`] = encoder.encode(
      typeof sa.color === "string" ? getColorXml(sa.color) : stringifyColorDefinitionPart(sa.color),
    );
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
  for (let i = 0; i < notesOptions.length; i++) {
    files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = encoder.encode(
      XML_DECL + (notesSlideDesc.stringify(notesOptions[i]!, descCtx) ?? ""),
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
  for (const [i, syncOpts] of slideSyncOptionsList.entries()) {
    files[`ppt/slideSyncPr/slideSyncPr${i + 1}.xml`] = encoder.encode(
      XML_DECL + (slideSyncDesc.stringify(syncOpts, descCtx) ?? ""),
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
    addBinaryFile(files, `ppt/media/${image.fileName}`, image.data, mediaLevel);
    if (image.type === "svg" && "fallback" in image) {
      const fallback = (
        image as MediaData & {
          fallback: { fileName: string; data: Uint8Array };
        }
      ).fallback;
      addBinaryFile(files, `ppt/media/${fallback.fileName}`, fallback.data, mediaLevel);
    }
  }

  // OLE embedding binaries (ppt/embeddings/oleObjectN.bin)
  for (const embedding of descCtx.embeddings) {
    addBinaryFile(files, `ppt/embeddings/${embedding.fileName}`, embedding.data, mediaLevel);
  }

  // Raw passthrough parts (handout masters, customXml, unknown extensions, …).
  // The compiler output above wins over a passthrough copy at the same path —
  // media/charts/notes absorbed into the model are re-emitted under pinned
  // source paths, so only what the model missed actually passes through.
  const passthroughSkipped = new Set<string>();
  for (const part of options.rawParts ?? []) {
    if (files[part.path] !== undefined) {
      passthroughSkipped.add(part.path);
      continue;
    }
    files[part.path] = toUint8Array(part.data);
  }

  // Derive [Content_Types].xml from the actual parts written — the file set is
  // the single source of truth, so declarations cannot drift from what is on
  // disk, and sparse/index-based names (slide-keyed comments) are handled
  // naturally because emission follows the files.
  const contentTypesInput = deriveContentTypes(Object.keys(files), {
    resolve: PPTX_CONTENT_TYPE_RESOLVER,
    mediaContentTypes: PPTX_MEDIA_CONTENT_TYPES,
    // Round-trip: the source declaration table is the base; derived entries
    // only fill what surviving source entries leave uncovered or mistyped.
    source: options.contentTypes,
    verbatimPaths: new Set((options.rawParts ?? []).map((p) => p.path)),
  });
  // Passthrough parts whose extension has no covering Default would leave the
  // package invalid (an undeclared part — Office refuses to open). Only those
  // borrow their source content-type declaration as a per-part Override;
  // extensions already covered (xml/rels/media) stay as derived above.
  const coveredExt = new Set(contentTypesInput.defaults.map((d) => d.extension.toLowerCase()));
  for (const part of options.rawParts ?? []) {
    if (part.contentType === undefined || passthroughSkipped.has(part.path)) continue;
    const dot = part.path.lastIndexOf(".");
    const slash = part.path.lastIndexOf("/");
    const ext = dot > slash ? part.path.slice(dot + 1).toLowerCase() : undefined;
    if (ext && coveredExt.has(ext)) continue;
    contentTypesInput.overrides.push({ partName: `/${part.path}`, contentType: part.contentType });
  }
  files["[Content_Types].xml"] = encoder.encode(
    XML_DECL + (contentTypesDesc.stringify(contentTypesInput, descCtx) ?? ""),
  );

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
