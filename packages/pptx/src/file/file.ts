import { AppProperties } from "@file/app-properties/app-properties";
import type { Background } from "@file/background/background";
import { ChartCollection } from "@file/chart/chart-collection";
import { ContentTypes } from "@file/content-types/content-types";
import { CoreProperties, type ICorePropertiesOptions } from "@file/core-properties/properties";
import type { IHeaderFooterOptions } from "@file/header-footer/header-footer";
import { HyperlinkCollection } from "@file/hyperlink-collection";
import { Media } from "@file/media/media";
import { NotesSlide } from "@file/notes/notes-slide";
import { PresentationProperties } from "@file/presentation-properties";
import { PresentationWrapper } from "@file/presentation/presentation-wrapper";
import { Relationships } from "@file/relationships/relationships";
import { DefaultSlideLayout } from "@file/slide-layout/slide-layout";
import { DefaultSlideMaster } from "@file/slide-master/slide-master";
import { Slide } from "@file/slide/slide";
import { SmartArtCollection } from "@file/smartart/smartart-collection";
import { TableStyles } from "@file/table-styles";
import { DefaultTheme } from "@file/theme/theme";
import type { ITransitionOptions } from "@file/transition/transition";
import { ViewProperties } from "@file/view-properties";
import type { BaseXmlComponent } from "@file/xml-components";
import type { RelationshipType } from "@office-open/core";
import { pixelsToEmus } from "@util/types";

export interface ISlideOptions {
    readonly children?: readonly BaseXmlComponent[];
    readonly background?: Background;
    readonly notes?: string;
    readonly transition?: ITransitionOptions;
    readonly headerFooter?: IHeaderFooterOptions;
}

export interface IShowOptions {
    readonly loop?: boolean;
    readonly kiosk?: boolean;
    readonly showNarration?: boolean;
    readonly useTimings?: boolean;
}

export interface IPresentationOptions extends ICorePropertiesOptions {
    readonly slideWidth?: number;
    readonly slideHeight?: number;
    readonly slides?: readonly ISlideOptions[];
    readonly show?: IShowOptions;
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

export class File {
    private readonly slideOptions: readonly ISlideOptions[];
    private readonly corePropsOptions: ICorePropertiesOptions;
    private readonly showOptions?: IShowOptions;
    private readonly slideWidthEmus?: number;
    private readonly slideHeightEmus?: number;

    // Lazy components — created on first getter access
    private coreProperties?: CoreProperties;
    private appProperties?: AppProperties;
    private contentTypes?: ContentTypes;
    private media?: Media;
    private charts?: ChartCollection;
    private smartArts?: SmartArtCollection;
    private hyperlinks?: HyperlinkCollection;
    private presentationWrapper?: PresentationWrapper;
    private theme?: DefaultTheme;
    private tableStyles?: TableStyles;
    private presProps?: PresentationProperties;
    private viewProps?: ViewProperties;
    private slideMaster?: DefaultSlideMaster;
    private slideLayout?: DefaultSlideLayout;
    private slideMasterRels?: Relationships;
    private slideLayoutRels?: Relationships;
    private notesMasterRels?: Relationships;

    // Lazy slide data — built on first access
    private slides?: Slide[];
    private slideWrappers?: Array<{ readonly View: Slide; readonly Relationships: Relationships }>;
    private notesSlides?: NotesSlide[];

    // Lazy relationship data — built on first access
    private fileRels?: Relationships;

    public constructor(options: IPresentationOptions) {
        const slides = options.slides ?? [];
        this.slideOptions = slides;
        this.corePropsOptions = options;
        this.showOptions = options.show;
        this.slideWidthEmus = options.slideWidth ? pixelsToEmus(options.slideWidth) : undefined;
        this.slideHeightEmus = options.slideHeight ? pixelsToEmus(options.slideHeight) : undefined;
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
            for (let i = 0; i < this.slideOptions.length; i++) {
                this.contentTypes.addSlide(i + 1);
                if (this.slideOptions[i].notes) {
                    this.contentTypes.addNotesSlide(i + 1);
                }
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
            this.presentationWrapper = new PresentationWrapper({
                slideWidth: this.slideWidthEmus,
                slideHeight: this.slideHeightEmus,
                slideIds: this.slideOptions.map((_, i) => 256 + i),
            });
            // Presentation-level relationships
            const presRels = this.PresentationWrapper.Relationships;
            presRels.addRelationship(
                1,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
                "slideMasters/slideMaster1.xml",
            );
            for (let i = 0; i < this.slideOptions.length; i++) {
                presRels.addRelationship(
                    i + 2,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                    `slides/slide${i + 1}.xml`,
                );
            }
            const n = this.slideOptions.length + 2;
            presRels.addRelationship(
                n,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
                "presProps.xml",
            );
            presRels.addRelationship(
                n + 1,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
                "viewProps.xml",
            );
            presRels.addRelationship(
                n + 2,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                "theme/theme1.xml",
            );
            presRels.addRelationship(
                n + 3,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
                "tableStyles.xml",
            );
        }
        return this.presentationWrapper;
    }

    public get Theme(): DefaultTheme {
        return (this.theme ??= new DefaultTheme());
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

    public get SlideMaster(): DefaultSlideMaster {
        if (!this.slideMaster) {
            const hf = this.slideOptions.find((s) => s.headerFooter)?.headerFooter;
            this.slideMaster = new DefaultSlideMaster(hf);
        }
        return this.slideMaster;
    }

    public get SlideMasterRelationships(): Relationships {
        if (!this.slideMasterRels) {
            this.slideMasterRels = buildRelationships([
                {
                    id: 1,
                    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                    target: "../slideLayouts/slideLayout1.xml",
                },
                {
                    id: 2,
                    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                    target: "../theme/theme1.xml",
                },
            ]);
        }
        return this.slideMasterRels;
    }

    public get SlideLayout(): DefaultSlideLayout {
        return (this.slideLayout ??= new DefaultSlideLayout());
    }

    public get SlideLayoutRelationships(): Relationships {
        if (!this.slideLayoutRels) {
            this.slideLayoutRels = buildRelationships([
                {
                    id: 1,
                    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
                    target: "../slideMasters/slideMaster1.xml",
                },
            ]);
        }
        return this.slideLayoutRels;
    }

    public get Slides(): readonly Slide[] {
        if (!this.slides) {
            this.slides = [];
            for (const s of this.slideOptions) {
                this.slides.push(
                    new Slide(s.children ?? [], s.background, s.transition, s.headerFooter),
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
                this.slideWrappers.push({
                    View: this.Slides[i],
                    Relationships: buildRelationships([
                        {
                            id: 1,
                            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                            target: "../slideLayouts/slideLayout1.xml",
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

    public get NotesMasterRelationships(): Relationships {
        return (this.notesMasterRels ??= new Relationships());
    }
}
