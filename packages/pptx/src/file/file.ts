import { AppProperties } from "@file/app-properties/app-properties";
import { Background } from "@file/background/background";
import { ChartCollection } from "@file/chart/chart-collection";
import { ContentTypes } from "@file/content-types/content-types";
import { CoreProperties, type ICorePropertiesOptions } from "@file/core-properties/properties";
import { Media } from "@file/media/media";
import { PresentationProperties } from "@file/presentation-properties";
import { PresentationWrapper } from "@file/presentation/presentation-wrapper";
import { Relationships } from "@file/relationships/relationships";
import { DefaultSlideLayout } from "@file/slide-layout/slide-layout";
import { DefaultSlideMaster } from "@file/slide-master/slide-master";
import { Slide } from "@file/slide/slide";
import { TableStyles } from "@file/table-styles";
import { DefaultTheme } from "@file/theme/theme";
import { ViewProperties } from "@file/view-properties";
import type { XmlComponent } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

export interface ISlideOptions {
    readonly children?: readonly XmlComponent[];
    readonly background?: Background;
}

export interface IPresentationOptions extends ICorePropertiesOptions {
    readonly slideWidth?: number;
    readonly slideHeight?: number;
    readonly slides?: readonly ISlideOptions[];
}

export class File {
    private readonly coreProperties: CoreProperties;
    private readonly appProperties: AppProperties;
    private readonly contentTypes: ContentTypes;
    private readonly fileRelationships: Relationships;
    private readonly media: Media;
    private readonly charts: ChartCollection;
    private readonly presentationWrapper: PresentationWrapper;
    private readonly theme: DefaultTheme;
    private readonly tableStyles: TableStyles;
    private readonly presProps: PresentationProperties;
    private readonly viewProps: ViewProperties;
    private readonly slideMaster: DefaultSlideMaster;
    private readonly slideLayout: DefaultSlideLayout;
    private readonly slideMasterRelationships: Relationships;
    private readonly slideLayoutRelationships: Relationships;
    private readonly slides: Slide[];
    private readonly slideWrappers: Array<{
        readonly View: Slide;
        readonly Relationships: Relationships;
    }>;

    public constructor(options: IPresentationOptions) {
        const slides = options.slides ?? [];

        this.coreProperties = new CoreProperties(options);
        this.appProperties = new AppProperties();
        this.contentTypes = new ContentTypes();
        this.media = new Media();
        this.charts = new ChartCollection();

        // Package-level relationships (_rels/.rels)
        this.fileRelationships = new Relationships();
        this.fileRelationships.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "ppt/presentation.xml",
        );
        this.fileRelationships.addRelationship(
            2,
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "docProps/core.xml",
        );
        this.fileRelationships.addRelationship(
            3,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
            "docProps/app.xml",
        );

        // Create slides and their wrappers
        this.slides = [];
        this.slideWrappers = [];
        for (let i = 0; i < slides.length; i++) {
            const slide = new Slide(slides[i].children ?? [], slides[i].background);
            this.slides.push(slide);
            const slideRels = new Relationships();
            slideRels.addRelationship(
                1,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                "../slideLayouts/slideLayout1.xml",
            );
            this.slideWrappers.push({ View: slide, Relationships: slideRels });
            this.contentTypes.addSlide(i + 1);
        }

        // Presentation wrapper with slide ID list
        this.presentationWrapper = new PresentationWrapper({
            slideWidth: options.slideWidth ? pixelsToEmus(options.slideWidth) : undefined,
            slideHeight: options.slideHeight ? pixelsToEmus(options.slideHeight) : undefined,
            slideIds: slides.map((_, i) => 256 + i),
        });

        // Presentation-level relationships (ppt/_rels/presentation.xml.rels)
        this.presentationWrapper.Relationships.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
            "slideMasters/slideMaster1.xml",
        );
        for (let i = 0; i < slides.length; i++) {
            this.presentationWrapper.Relationships.addRelationship(
                i + 2,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                `slides/slide${i + 1}.xml`,
            );
        }

        const nextRId = slides.length + 2;
        this.presentationWrapper.Relationships.addRelationship(
            nextRId,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
            "presProps.xml",
        );
        this.presentationWrapper.Relationships.addRelationship(
            nextRId + 1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
            "viewProps.xml",
        );
        this.presentationWrapper.Relationships.addRelationship(
            nextRId + 2,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            "theme/theme1.xml",
        );
        this.presentationWrapper.Relationships.addRelationship(
            nextRId + 3,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
            "tableStyles.xml",
        );

        // Theme
        this.theme = new DefaultTheme();

        // Table styles (required when presentation contains tables)
        this.tableStyles = new TableStyles();
        this.presProps = new PresentationProperties();
        this.viewProps = new ViewProperties();

        // Slide Master
        this.slideMaster = new DefaultSlideMaster();
        this.slideMasterRelationships = new Relationships();
        this.slideMasterRelationships.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
            "../slideLayouts/slideLayout1.xml",
        );
        this.slideMasterRelationships.addRelationship(
            2,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            "../theme/theme1.xml",
        );

        // Slide Layout
        this.slideLayout = new DefaultSlideLayout();
        this.slideLayoutRelationships = new Relationships();
        this.slideLayoutRelationships.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
            "../slideMasters/slideMaster1.xml",
        );
    }

    public get CoreProperties(): CoreProperties {
        return this.coreProperties;
    }

    public get AppProperties(): AppProperties {
        return this.appProperties;
    }

    public get ContentTypes(): ContentTypes {
        return this.contentTypes;
    }

    public get FileRelationships(): Relationships {
        return this.fileRelationships;
    }

    public get Media(): Media {
        return this.media;
    }

    public get Charts(): ChartCollection {
        return this.charts;
    }

    public get PresentationWrapper(): PresentationWrapper {
        return this.presentationWrapper;
    }

    public get Theme(): DefaultTheme {
        return this.theme;
    }

    public get TableStyles(): TableStyles {
        return this.tableStyles;
    }

    public get PresProps(): PresentationProperties {
        return this.presProps;
    }

    public get ViewProps(): ViewProperties {
        return this.viewProps;
    }

    public get SlideMaster(): DefaultSlideMaster {
        return this.slideMaster;
    }

    public get SlideMasterRelationships(): Relationships {
        return this.slideMasterRelationships;
    }

    public get SlideLayout(): DefaultSlideLayout {
        return this.slideLayout;
    }

    public get SlideLayoutRelationships(): Relationships {
        return this.slideLayoutRelationships;
    }

    public get Slides(): readonly Slide[] {
        return this.slides;
    }

    public get SlideWrappers(): Array<{
        readonly View: Slide;
        readonly Relationships: Relationships;
    }> {
        return this.slideWrappers;
    }
}
