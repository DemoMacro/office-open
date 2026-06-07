import type { AnimationOptions } from "@file/animation/types";
import type { BackgroundOptions } from "@file/background";
import type { AuthorEntry, CommentEntry } from "@file/comment";
import { ContentTypes } from "@file/content-types";
import type { CorePropertiesOptions } from "@file/core-properties";
import type { HandoutMasterOptions } from "@file/handout-master";
import type { SlideHeaderFooterOptions } from "@file/header-footer";
import { HyperlinkCollection } from "@file/hyperlink-collection";
import { Media } from "@file/media/media";
import type { NotesMasterOptions } from "@file/notes-master";
import type { PresentationOptions as PresInternalOptions } from "@file/presentation";
import type {
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
} from "@file/presentation";
import type { PresentationPropertiesFullOptions } from "@file/presentation-properties";
import type {
  WebPropertiesOptions,
  PrintPropertiesOptions,
  HtmlPublishPropertiesOptions,
} from "@file/presentation-properties";
import type { ShapeOptions } from "@file/shape/shape";
import { SlideLayout, type SlideLayoutType } from "@file/slide-layout";
import {
  DefaultSlideMaster,
  type MasterPlaceholderPosition,
  type SlideMasterOptions,
} from "@file/slide-master";
import type { ControlOptions } from "@file/slide/slide";
import type { SlideChild } from "@file/slide/slide-child";
import type { SlideSyncOptions } from "@file/slide/slide-sync-properties";
import { DefaultTheme, type ThemeOptions } from "@file/theme";
import type { TransitionOptions } from "@file/transition";
import type { ViewPropertiesOptions } from "@file/view-properties";
import type { BaseXmlComponent } from "@file/xml-components";
import { Relationships } from "@office-open/core";
import type { RelationshipType } from "@office-open/core";
import { convertPixelsToEmu } from "@office-open/core";
import type { TableStyleListOptions } from "@office-open/core";
import { ChartCollection } from "@office-open/core/chart";
import { SmartArtCollection } from "@office-open/core/smartart";

// ── Public interfaces ──

export type MasterChild = BaseXmlComponent | { shape: ShapeOptions };

export type SlideSize = "16:9" | "4:3" | { readonly width: number; readonly height: number };

export interface LayoutPlaceholderOptions {
  readonly title?: MasterPlaceholderPosition | false;
  readonly body?: MasterPlaceholderPosition | false;
  readonly subtitle?: MasterPlaceholderPosition | false;
  readonly date?: MasterPlaceholderPosition | false;
  readonly footer?: MasterPlaceholderPosition | false;
  readonly slideNumber?: MasterPlaceholderPosition | false;
}

export interface LayoutDefinition {
  readonly type?: SlideLayoutType;
  readonly name?: string;
  readonly matchingName?: string;
  readonly placeholders?: LayoutPlaceholderOptions;
  readonly children?: readonly MasterChild[];
}

export interface MasterDefinition extends SlideMasterOptions {
  readonly name?: string;
  readonly theme?: ThemeOptions;
  readonly layouts?: readonly LayoutDefinition[];
}

export interface SlideCommentOptions {
  readonly author: string;
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly initials?: string;
  readonly date?: string;
}

export interface SlideAnimation {
  readonly shapeId: number;
  readonly options: AnimationOptions;
}

export interface SlideOptions {
  readonly children?: readonly SlideChild[];
  readonly background?: BackgroundOptions;
  readonly notes?: string;
  readonly transition?: TransitionOptions;
  readonly headerFooter?: SlideHeaderFooterOptions;
  readonly comments?: readonly SlideCommentOptions[];
  readonly layout?: SlideLayoutType | string;
  readonly master?: string;
  readonly showMasterShapes?: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly controls?: readonly ControlOptions[];
  readonly customerData?: readonly { readonly rId: string }[];
  readonly slideSync?: SlideSyncOptions;
  readonly animations?: readonly SlideAnimation[];
}

export interface ShowOptions {
  readonly loop?: boolean;
  /** Show type: "present" (default), "browse", or "kiosk" */
  readonly type?: "present" | "browse" | "kiosk";
  /** Browse mode: show scrollbar (default true) */
  readonly showScrollbar?: boolean;
  /** Kiosk mode: restart interval in ms (default 300000) */
  readonly restart?: number;
  readonly showNarration?: boolean;
  readonly showAnimation?: boolean;
  readonly useTimings?: boolean;
  /** Slide range for show (1-based start/end) */
  readonly slideRange?: { readonly start: number; readonly end: number };
  /** Pen color (hex ARGB, e.g. "FFFF0000" for red) */
  readonly penColor?: string;
}

export interface PresentationOptions extends CorePropertiesOptions {
  readonly size?: SlideSize;
  readonly masters?: readonly MasterDefinition[];
  readonly slides?: readonly SlideOptions[];
  readonly show?: ShowOptions;
  readonly view?: ViewPropertiesOptions;
  readonly includeHandoutMaster?: boolean;
  /** Include notes master (auto-enabled when slides have notes) */
  readonly includeNotesMaster?: boolean;
  /** Handout master customization options */
  readonly handoutMasterOptions?: HandoutMasterOptions;
  /** Notes master customization options */
  readonly notesMasterOptions?: NotesMasterOptions;
  readonly tableStyles?: TableStyleListOptions;
  readonly web?: WebPropertiesOptions;
  readonly print?: PrintPropertiesOptions;
  readonly htmlPublish?: HtmlPublishPropertiesOptions;
  readonly serverZoom?: string;
  readonly firstSlideNum?: number;
  readonly showSpecialPlsOnTitleSld?: boolean;
  readonly rtl?: boolean;
  readonly removePersonalInfoOnSave?: boolean;
  readonly compatMode?: boolean;
  readonly strictFirstAndLastChars?: boolean;
  readonly embedTrueTypeFonts?: boolean;
  readonly saveSubsetFonts?: boolean;
  readonly autoCompressPictures?: boolean;
  readonly bookmarkIdSeed?: number;
  readonly conformance?: "strict" | "transitional";
  readonly photoAlbum?: PhotoAlbumOptions;
  readonly modifyVerifier?: ModifyVerifierOptions;
  readonly embeddedFonts?: EmbeddedFontOptions[];
  readonly customShows?: CustomShowOptions[];
  readonly kinsoku?: KinsokuOptions[];
  readonly customerData?: CustomerDataOptions;
  readonly colorMru?: readonly string[];
}

interface RelEntry {
  readonly id: number | string;
  readonly type: RelationshipType;
  readonly target: string;
  readonly mode?: string;
}

function buildRelationships(entries: readonly RelEntry[]): Relationships {
  const rels = new Relationships();
  for (const e of entries) {
    rels.addRelationship(e.id, e.type, e.target, e.mode as "External" | undefined);
  }
  return rels;
}

function deriveInitials(name: string): string {
  const parts = name.trim().split(/\s+/);
  return parts.length >= 2
    ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
    : name.slice(0, 2).toUpperCase();
}

function resolveSlideSize(size?: SlideSize): { width: number; height: number } {
  if (!size || size === "16:9") return { width: 12192000, height: 6858000 };
  if (size === "4:3") return { width: 9144000, height: 6858000 };
  return { width: convertPixelsToEmu(size.width), height: convertPixelsToEmu(size.height) };
}

interface LayoutInfo {
  readonly key: string;
  readonly index: number;
  readonly masterIndex: number;
  readonly layout: SlideLayout;
}

interface MasterInfo {
  readonly name: string;
  readonly index: number;
  readonly definition: MasterDefinition;
  readonly master: DefaultSlideMaster;
  readonly theme: DefaultTheme;
  readonly layouts: LayoutInfo[];
  readonly masterRels: Relationships;
  readonly layoutRels: Relationships[];
}

export class File {
  /** @internal Accessible by descriptor-based compiler. */
  public readonly slideOptions: readonly SlideOptions[];
  /** @internal Raw core properties options (for descriptor pipeline). */
  public readonly corePropsOptions: CorePropertiesOptions;
  private readonly slideWidthEmus: number;
  private readonly slideHeightEmus: number;
  private readonly masterDefs: readonly MasterDefinition[];
  private readonly includeHandout: boolean;
  public readonly tableStylesOpts?: TableStyleListOptions;
  public readonly viewOpts?: ViewPropertiesOptions;
  private readonly handoutMasterOpts?: HandoutMasterOptions;
  private readonly notesMasterOpts?: NotesMasterOptions;
  public readonly presPropsFullOpts?: PresentationPropertiesFullOptions;
  private readonly presAttrOpts?: Partial<
    Pick<
      PresInternalOptions,
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
  >;

  // Lazy components
  private _contentTypes?: ContentTypes;
  private _media?: Media;
  private _charts?: ChartCollection;
  private _smartArts?: SmartArtCollection;
  private _hyperlinks?: HyperlinkCollection;
  private _presRels?: Relationships;
  private _presOptions?: PresInternalOptions;
  private _notesMasterRels?: Relationships;

  // Multi-master support
  private masterMap?: MasterInfo[];
  private _allLayouts?: readonly LayoutInfo[];
  private _allLayoutRels?: readonly Relationships[];

  // Lazy slide data
  private _slideRels?: Relationships[];
  private _notesTexts?: string[];
  private _commentAuthorEntries?: AuthorEntry[];
  private _slideCommentEntries?: (CommentEntry[] | undefined)[];
  private _slideSyncOptions?: SlideSyncOptions[];
  private _slideSyncIndexMap?: Map<number, number>;

  // Lazy relationship data
  private _fileRels?: Relationships;

  public constructor(options: PresentationOptions) {
    this.slideOptions = options.slides ?? [];
    this.corePropsOptions = options;
    this.masterDefs = options.masters ?? [];
    this.includeHandout = options.includeHandoutMaster ?? false;
    this.tableStylesOpts = options.tableStyles;
    this.viewOpts = options.view;
    this.presPropsFullOpts =
      options.web || options.print || options.htmlPublish || options.colorMru
        ? {
            web: options.web,
            print: options.print,
            htmlPublish: options.htmlPublish,
            colorMru: options.colorMru,
          }
        : undefined;
    this.handoutMasterOpts = options.handoutMasterOptions;
    this.notesMasterOpts = options.notesMasterOptions;
    const sz = resolveSlideSize(options.size);
    this.slideWidthEmus = sz.width;
    this.slideHeightEmus = sz.height;
    // Pass presentation attributes through to Presentation
    if (
      options.serverZoom ||
      options.firstSlideNum !== undefined ||
      options.showSpecialPlsOnTitleSld !== undefined ||
      options.rtl !== undefined ||
      options.removePersonalInfoOnSave !== undefined ||
      options.compatMode !== undefined ||
      options.strictFirstAndLastChars !== undefined ||
      options.embedTrueTypeFonts !== undefined ||
      options.saveSubsetFonts !== undefined ||
      options.autoCompressPictures !== undefined ||
      options.bookmarkIdSeed !== undefined ||
      options.conformance !== undefined ||
      options.photoAlbum !== undefined ||
      options.modifyVerifier !== undefined ||
      options.embeddedFonts !== undefined ||
      options.customShows !== undefined ||
      options.kinsoku !== undefined ||
      options.customerData !== undefined
    ) {
      this.presAttrOpts = {
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
  }

  // ── Master / Layout resolution ──

  private getMasterMap(): MasterInfo[] {
    if (this.masterMap) return this.masterMap;

    const defs = this.masterDefs.length > 0 ? this.masterDefs : [{} as MasterDefinition];
    const slideMasterLookup = new Map<number, number>();

    // Build slide → master index lookup
    for (let si = 0; si < this.slideOptions.length; si++) {
      const masterName = this.slideOptions[si].master;
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

      // Collect layout types needed for this master
      const layoutDefs = def.layouts;
      let layoutKeys: readonly string[];
      if (layoutDefs && layoutDefs.length > 0) {
        layoutKeys = layoutDefs.map(
          (ld) => ld.type ?? ld.name ?? `layout${mi}_${layoutDefs.indexOf(ld)}`,
        );
      } else {
        // Auto-derive from slides referencing this master
        const seen = new Set<string>();
        const keys: string[] = [];
        for (let si = 0; si < this.slideOptions.length; si++) {
          if (slideMasterLookup.get(si) === mi) {
            const lt = this.slideOptions[si].layout ?? "blank";
            if (!seen.has(lt)) {
              seen.add(lt);
              keys.push(lt);
            }
          }
        }
        layoutKeys = keys.length > 0 ? keys : ["blank"];
      }

      const hf = this.slideOptions.find((s) => s.headerFooter)?.headerFooter;
      const master = new DefaultSlideMaster(layoutKeys.length, hf, def, this.slideWidthEmus, mi);
      const theme = new DefaultTheme(def.theme);

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
          layout: new SlideLayout(slideLayoutType, this.slideWidthEmus, layoutDef),
        });
        layoutRels.push(
          buildRelationships([
            {
              id: 1,
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
              target: `../slideMasters/slideMaster${mi + 1}.xml`,
            },
          ]),
        );
        globalLayoutIndex++;
      }

      // Master rels: layouts + theme
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
        definition: def,
        master,
        theme,
        layouts,
        masterRels: buildRelationships(masterRelsEntries),
        layoutRels,
      });
    }

    this.masterMap = masters;
    this._allLayouts = masters.flatMap((m) => m.layouts);
    this._allLayoutRels = masters.flatMap((m) => m.layoutRels);
    return this.masterMap;
  }

  private findLayoutForSlide(slideIndex: number): LayoutInfo {
    const opts = this.slideOptions[slideIndex];
    const masters = this.getMasterMap();
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

  // ── Lazy getters ──

  public get contentTypes(): ContentTypes {
    if (!this._contentTypes) {
      this._contentTypes = new ContentTypes();
      let hasComments = false;
      let notesSlideIdx = 0;
      let slideSyncIdx = 0;
      for (let i = 0; i < this.slideOptions.length; i++) {
        this._contentTypes.addSlide(i + 1);
        if (this.slideOptions[i].notes) {
          this._contentTypes.addNotesSlide(notesSlideIdx + 1);
          notesSlideIdx++;
        }
        if (this.slideOptions[i].comments && this.slideOptions[i].comments!.length > 0) {
          this._contentTypes.addComments(i + 1);
          hasComments = true;
        }
        if (this.slideOptions[i].slideSync) {
          this._contentTypes.addSlideSyncPr(slideSyncIdx + 1);
          slideSyncIdx++;
        }
      }
      if (notesSlideIdx > 0) {
        this._contentTypes.addNotesMaster();
      }
      if (hasComments) {
        this._contentTypes.addCommentAuthors();
      }
      if (this.includeHandout) {
        this._contentTypes.addHandoutMaster();
      }
    }
    return this._contentTypes;
  }

  public get fileRelationships(): Relationships {
    if (!this._fileRels) {
      this._fileRels = buildRelationships([
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
      ]);
    }
    return this._fileRels;
  }

  public get media(): Media {
    return (this._media ??= new Media());
  }

  public get charts(): ChartCollection {
    return (this._charts ??= new ChartCollection());
  }

  public get smartArts(): SmartArtCollection {
    return (this._smartArts ??= new SmartArtCollection());
  }

  public get hyperlinks(): HyperlinkCollection {
    return (this._hyperlinks ??= new HyperlinkCollection());
  }

  /** Initialize presentation state (options + relationships). */
  private initPresentationState(): void {
    if (this._presOptions) return;
    const masters = this.getMasterMap();
    this._presOptions = {
      slideWidth: this.slideWidthEmus,
      slideHeight: this.slideHeightEmus,
      slideIds: this.slideOptions.map((_, i) => 256 + i),
      masterCount: masters.length,
      ...this.presAttrOpts,
    };
    this._presRels = new Relationships();
    let rid = 1;
    // Masters
    for (let mi = 0; mi < masters.length; mi++) {
      this._presRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
        `slideMasters/slideMaster${mi + 1}.xml`,
      );
    }
    // Slides
    for (let i = 0; i < this.slideOptions.length; i++) {
      this._presRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        `slides/slide${i + 1}.xml`,
      );
    }
    // Static parts
    this._presRels.addRelationship(
      rid++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
      "presProps.xml",
    );
    this._presRels.addRelationship(
      rid++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
      "viewProps.xml",
    );
    for (let mi = 0; mi < masters.length; mi++) {
      this._presRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        `theme/theme${mi + 1}.xml`,
      );
    }
    this._presRels.addRelationship(
      rid,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
      "tableStyles.xml",
    );
  }

  /** Mutable presentation options (for compiler to set notesMasterRId etc.) */
  public get presOptions(): PresInternalOptions {
    this.initPresentationState();
    return this._presOptions!;
  }

  /** Mutable presentation relationships. */
  public get presRels(): Relationships {
    this.initPresentationState();
    return this._presRels!;
  }

  public get themes(): readonly DefaultTheme[] {
    return this.getMasterMap().map((m) => m.theme);
  }

  public get slideMasters(): readonly DefaultSlideMaster[] {
    return this.getMasterMap().map((m) => m.master);
  }

  public get slideMasterRelsArray(): readonly Relationships[] {
    return this.getMasterMap().map((m) => m.masterRels);
  }

  public get allLayouts(): readonly LayoutInfo[] {
    this.getMasterMap();
    return this._allLayouts!;
  }

  public get allLayoutRelsArray(): readonly Relationships[] {
    this.getMasterMap();
    return this._allLayoutRels!;
  }

  public get slideRelationships(): readonly Relationships[] {
    if (!this._slideRels) {
      this._slideRels = [];
      for (let i = 0; i < this.slideOptions.length; i++) {
        const layout = this.findLayoutForSlide(i);
        this._slideRels.push(
          buildRelationships([
            {
              id: 1,
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
              target: `../slideLayouts/slideLayout${layout.index + 1}.xml`,
            },
          ]),
        );
      }
    }
    return this._slideRels;
  }

  /** Notes slide text data (for descriptor pipeline). */
  public get notesTexts(): readonly string[] {
    if (!this._notesTexts) {
      this._notesTexts = [];
      for (let i = 0; i < this.slideOptions.length; i++) {
        if (this.slideOptions[i].notes) {
          this._notesTexts.push(this.slideOptions[i].notes!);
        }
      }
    }
    return this._notesTexts;
  }

  /** Map from slide index → notesSlide index (0-based) for slides that have notes. */
  public get notesSlideIndexMap(): Map<number, number> {
    const map = new Map<number, number>();
    let notesIdx = 0;
    for (let i = 0; i < this.slideOptions.length; i++) {
      if (this.slideOptions[i].notes) {
        map.set(i, notesIdx++);
      }
    }
    return map;
  }

  public get notesMasterRelationships(): Relationships {
    return (this._notesMasterRels ??= new Relationships());
  }

  public get slideSyncOptionsList(): readonly SlideSyncOptions[] {
    if (!this._slideSyncOptions) {
      this._slideSyncOptions = [];
      for (const s of this.slideOptions) {
        if (s.slideSync) {
          this._slideSyncOptions.push(s.slideSync);
        }
      }
    }
    return this._slideSyncOptions;
  }

  /** Map from slide index → slideSyncPr index (0-based) for slides that have slideSync. */
  public get slideSyncIndexMap(): Map<number, number> {
    if (!this._slideSyncIndexMap) {
      this._slideSyncIndexMap = new Map();
      let syncIdx = 0;
      for (let i = 0; i < this.slideOptions.length; i++) {
        if (this.slideOptions[i].slideSync) {
          this._slideSyncIndexMap.set(i, syncIdx++);
        }
      }
    }
    return this._slideSyncIndexMap;
  }

  public get hasHandoutMaster(): boolean {
    return this.includeHandout;
  }

  public get hasOutlineViewSlides(): boolean {
    return !!this.viewOpts?.outlineView?.slides && this.viewOpts.outlineView.slides.length > 0;
  }

  public get slideCount(): number {
    return this.slideOptions.length;
  }

  /** HTML publish target info (for creating presProps.xml.rels) */
  public get htmlPublishInfo(): { readonly rId: string; readonly target?: string } | undefined {
    const hp = this.presPropsFullOpts?.htmlPublish;
    return hp?.rId ? { rId: hp.rId, target: hp.target } : undefined;
  }

  public get handoutMasterOptions(): HandoutMasterOptions | undefined {
    return this.handoutMasterOpts;
  }

  public get notesMasterOptions(): NotesMasterOptions | undefined {
    return this.notesMasterOpts;
  }

  /** Raw comment author entries (for descriptor pipeline). */
  public get commentAuthorEntries(): AuthorEntry[] | undefined {
    if (!this._commentAuthorEntries) {
      this.buildComments();
    }
    return this._commentAuthorEntries;
  }

  /** Raw per-slide comment entries (for descriptor pipeline). */
  public get slideCommentEntries(): (CommentEntry[] | undefined)[] {
    if (!this._slideCommentEntries) {
      this.buildComments();
    }
    return this._slideCommentEntries!;
  }

  private buildComments(): void {
    const authorMap = new Map<
      string,
      { id: number; name: string; initials: string; clrIdx: number; commentCount: number }
    >();
    let nextAuthorId = 0;

    this._slideCommentEntries = Array.from<CommentEntry[] | undefined>({
      length: this.slideOptions.length,
    });

    for (let i = 0; i < this.slideOptions.length; i++) {
      const slideComments = this.slideOptions[i].comments;
      if (!slideComments || slideComments.length === 0) continue;

      const commentEntries: Array<{
        readonly authorId: number;
        readonly idx: number;
        readonly date?: string;
        readonly x: number;
        readonly y: number;
        readonly text: string;
      }> = [];

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
        });
      }

      this._slideCommentEntries[i] = commentEntries;
    }

    if (authorMap.size > 0) {
      const authors: AuthorEntry[] = Array.from(authorMap.values(), (a) => ({
        id: a.id,
        name: a.name,
        initials: a.initials,
        clrIdx: a.clrIdx,
        lastIdx: a.commentCount,
      }));
      this._commentAuthorEntries = authors;
    }
  }
}
