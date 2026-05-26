import { AppProperties } from "@file/app-properties/app-properties";
import { Background, type BackgroundOptions } from "@file/background/background";
import { ChartCollection } from "@file/chart/chart-collection";
import { CommentAuthorList } from "@file/comment/comment-author-list";
import type { AuthorEntry } from "@file/comment/comment-author-list";
import { SlideCommentList } from "@file/comment/slide-comment-list";
import { ContentTypes } from "@file/content-types/content-types";
import { CoreProperties, type CorePropertiesOptions } from "@file/core-properties/properties";
import type { SlideHeaderFooterOptions } from "@file/header-footer/header-footer";
import { HyperlinkCollection } from "@file/hyperlink-collection";
import { Media } from "@file/media/media";
import { NotesSlide } from "@file/notes/notes-slide";
import { PresentationProperties } from "@file/presentation-properties";
import { PresentationWrapper } from "@file/presentation/presentation-wrapper";
import { Relationships } from "@file/relationships/relationships";
import type { ShapeOptions } from "@file/shape/shape";
import { SlideLayout, type SlideLayoutType } from "@file/slide-layout/slide-layout";
import {
  DefaultSlideMaster,
  type MasterPlaceholderPosition,
  type SlideMasterOptions,
} from "@file/slide-master/slide-master";
import { Slide } from "@file/slide/slide";
import type { SlideChild } from "@file/slide/slide-child";
import { SmartArtCollection } from "@file/smartart/smartart-collection";
import { TableStyles } from "@file/table-styles";
import { DefaultTheme, type ThemeOptions } from "@file/theme/theme";
import type { TransitionOptions } from "@file/transition/transition";
import { ViewProperties } from "@file/view-properties";
import type { BaseXmlComponent } from "@file/xml-components";
import type { RelationshipType } from "@office-open/core";
import { convertPixelsToEmu } from "@office-open/core";

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

export interface SlideOptions {
  readonly children?: readonly SlideChild[];
  readonly background?: BackgroundOptions;
  readonly notes?: string;
  readonly transition?: TransitionOptions;
  readonly headerFooter?: SlideHeaderFooterOptions;
  readonly comments?: readonly SlideCommentOptions[];
  readonly layout?: SlideLayoutType | string;
  readonly master?: string;
}

export interface ShowOptions {
  readonly loop?: boolean;
  readonly kiosk?: boolean;
  readonly showNarration?: boolean;
  readonly useTimings?: boolean;
}

export interface PresentationOptions extends CorePropertiesOptions {
  readonly size?: SlideSize;
  readonly masters?: readonly MasterDefinition[];
  readonly slides?: readonly SlideOptions[];
  readonly show?: ShowOptions;
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
  private readonly slideOptions: readonly SlideOptions[];
  private readonly corePropsOptions: CorePropertiesOptions;
  private readonly showOptions?: ShowOptions;
  private readonly slideWidthEmus: number;
  private readonly slideHeightEmus: number;
  private readonly masterDefs: readonly MasterDefinition[];

  // Lazy components
  private coreProperties?: CoreProperties;
  private appProperties?: AppProperties;
  private contentTypes?: ContentTypes;
  private media?: Media;
  private charts?: ChartCollection;
  private smartArts?: SmartArtCollection;
  private hyperlinks?: HyperlinkCollection;
  private presentationWrapper?: PresentationWrapper;
  private tableStyles?: TableStyles;
  private presProps?: PresentationProperties;
  private viewProps?: ViewProperties;
  private notesMasterRels?: Relationships;

  // Multi-master support
  private masterMap?: MasterInfo[];
  private allLayouts?: readonly LayoutInfo[];
  private allLayoutRels?: readonly Relationships[];

  // Lazy slide data
  private slides?: Slide[];
  private slideWrappers?: Array<{ readonly View: Slide; readonly Relationships: Relationships }>;
  private notesSlides?: NotesSlide[];
  private commentAuthorList?: CommentAuthorList;
  private slideCommentLists?: (SlideCommentList | undefined)[];

  // Lazy relationship data
  private fileRels?: Relationships;

  public constructor(options: PresentationOptions) {
    this.slideOptions = options.slides ?? [];
    this.corePropsOptions = options;
    this.showOptions = options.show;
    this.masterDefs = options.masters ?? [];
    const sz = resolveSlideSize(options.size);
    this.slideWidthEmus = sz.width;
    this.slideHeightEmus = sz.height;
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
    this.allLayouts = masters.flatMap((m) => m.layouts);
    this.allLayoutRels = masters.flatMap((m) => m.layoutRels);
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

  public get CoreProperties(): CoreProperties {
    return (this.coreProperties ??= new CoreProperties(this.corePropsOptions));
  }

  public get AppProperties(): AppProperties {
    return (this.appProperties ??= new AppProperties());
  }

  public get ContentTypes(): ContentTypes {
    if (!this.contentTypes) {
      this.contentTypes = new ContentTypes();
      let hasComments = false;
      for (let i = 0; i < this.slideOptions.length; i++) {
        this.contentTypes.addSlide(i + 1);
        if (this.slideOptions[i].notes) {
          this.contentTypes.addNotesSlide(i + 1);
        }
        if (this.slideOptions[i].comments && this.slideOptions[i].comments!.length > 0) {
          this.contentTypes.addComments(i + 1);
          hasComments = true;
        }
      }
      if (hasComments) {
        this.contentTypes.addCommentAuthors();
      }
    }
    return this.contentTypes;
  }

  public get FileRelationships(): Relationships {
    if (!this.fileRels) {
      this.fileRels = buildRelationships([
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
    return this.fileRels;
  }

  public get Media(): Media {
    return (this.media ??= new Media());
  }

  public get Charts(): ChartCollection {
    return (this.charts ??= new ChartCollection());
  }

  public get SmartArts(): SmartArtCollection {
    return (this.smartArts ??= new SmartArtCollection());
  }

  public get Hyperlinks(): HyperlinkCollection {
    return (this.hyperlinks ??= new HyperlinkCollection());
  }

  public get PresentationWrapper(): PresentationWrapper {
    if (!this.presentationWrapper) {
      const masters = this.getMasterMap();
      this.presentationWrapper = new PresentationWrapper({
        slideWidth: this.slideWidthEmus,
        slideHeight: this.slideHeightEmus,
        slideIds: this.slideOptions.map((_, i) => 256 + i),
        masterCount: masters.length,
      });
      const presRels = this.PresentationWrapper.Relationships;
      let rid = 1;
      // Masters
      for (let mi = 0; mi < masters.length; mi++) {
        presRels.addRelationship(
          rid++,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
          `slideMasters/slideMaster${mi + 1}.xml`,
        );
      }
      // Slides
      for (let i = 0; i < this.slideOptions.length; i++) {
        presRels.addRelationship(
          rid++,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
          `slides/slide${i + 1}.xml`,
        );
      }
      // Static parts
      presRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
        "presProps.xml",
      );
      presRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
        "viewProps.xml",
      );
      for (let mi = 0; mi < masters.length; mi++) {
        presRels.addRelationship(
          rid++,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
          `theme/theme${mi + 1}.xml`,
        );
      }
      presRels.addRelationship(
        rid,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
        "tableStyles.xml",
      );
    }
    return this.presentationWrapper;
  }

  public get Themes(): readonly DefaultTheme[] {
    return this.getMasterMap().map((m) => m.theme);
  }

  public get TableStyles(): TableStyles {
    return (this.tableStyles ??= new TableStyles());
  }

  public get PresProps(): PresentationProperties {
    return (this.presProps ??= new PresentationProperties(this.showOptions));
  }

  public get ViewProps(): ViewProperties {
    return (this.viewProps ??= new ViewProperties());
  }

  public get SlideMasters(): readonly DefaultSlideMaster[] {
    return this.getMasterMap().map((m) => m.master);
  }

  public get SlideMasterRelsArray(): readonly Relationships[] {
    return this.getMasterMap().map((m) => m.masterRels);
  }

  public get AllLayouts(): readonly LayoutInfo[] {
    this.getMasterMap();
    return this.allLayouts!;
  }

  public get AllLayoutRelsArray(): readonly Relationships[] {
    this.getMasterMap();
    return this.allLayoutRels!;
  }

  public get Slides(): readonly Slide[] {
    if (!this.slides) {
      this.slides = [];
      for (const s of this.slideOptions) {
        this.slides.push(
          new Slide(
            s.children ?? [],
            s.background ? new Background(s.background) : undefined,
            s.transition,
            s.headerFooter,
          ),
        );
      }
    }
    return this.slides;
  }

  public get SlideWrappers(): Array<{
    readonly View: Slide;
    readonly Relationships: Relationships;
  }> {
    if (!this.slideWrappers) {
      this.slideWrappers = [];
      for (let i = 0; i < this.slideOptions.length; i++) {
        const layout = this.findLayoutForSlide(i);
        this.slideWrappers.push({
          View: this.Slides[i],
          Relationships: buildRelationships([
            {
              id: 1,
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
              target: `../slideLayouts/slideLayout${layout.index + 1}.xml`,
            },
          ]),
        });
      }
    }
    return this.slideWrappers;
  }

  public get NotesSlides(): readonly NotesSlide[] {
    if (!this.notesSlides) {
      this.notesSlides = [];
      for (let i = 0; i < this.slideOptions.length; i++) {
        if (this.slideOptions[i].notes) {
          this.notesSlides.push(new NotesSlide({ text: this.slideOptions[i].notes }));
        }
      }
    }
    return this.notesSlides;
  }

  /** Map from slide index → notesSlide index (0-based) for slides that have notes. */
  public get NotesSlideIndexMap(): Map<number, number> {
    const map = new Map<number, number>();
    let notesIdx = 0;
    for (let i = 0; i < this.slideOptions.length; i++) {
      if (this.slideOptions[i].notes) {
        map.set(i, notesIdx++);
      }
    }
    return map;
  }

  public get NotesMasterRelationships(): Relationships {
    return (this.notesMasterRels ??= new Relationships());
  }

  public get CommentAuthorList(): CommentAuthorList | undefined {
    if (!this.commentAuthorList && !this.slideCommentLists) {
      this.buildComments();
    }
    return this.commentAuthorList;
  }

  public get SlideCommentLists(): readonly (SlideCommentList | undefined)[] {
    if (!this.slideCommentLists) {
      this.buildComments();
    }
    return this.slideCommentLists!;
  }

  private buildComments(): void {
    const authorMap = new Map<
      string,
      { id: number; name: string; initials: string; clrIdx: number; commentCount: number }
    >();
    let nextAuthorId = 0;

    this.slideCommentLists = Array.from<SlideCommentList | undefined>({
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

      this.slideCommentLists[i] = new SlideCommentList(commentEntries);
    }

    if (authorMap.size > 0) {
      const authors: AuthorEntry[] = [...authorMap.values()].map((a) => ({
        id: a.id,
        name: a.name,
        initials: a.initials,
        clrIdx: a.clrIdx,
        lastIdx: a.commentCount,
      }));
      this.commentAuthorList = new CommentAuthorList(authors);
    }
  }
}
