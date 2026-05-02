import { DefaultNotesMaster } from "@file/notes-master/notes-master";
import { Formatter } from "@export/formatter";
import type { ChartCollection } from "@file/chart/chart-collection";
import type { File } from "@file/file";
import type { IContext } from "@file/xml-components";
import { xml } from "@office-open/xml";
import type { Zippable } from "fflate";

import { ChartReplacer } from "./chart-replacer";
import { HyperlinkReplacer } from "./hyperlink-replacer";
import { ImageReplacer } from "./image-replacer";
import { MediaReplacer } from "./media-replacer";

export interface IXmlifyedFile {
    readonly data: string | Uint8Array;
    readonly path: string;
}

export interface IXmlifyedFileMapping {
    [key: string]: IXmlifyedFile;
}

export class Compiler {
    private readonly formatter = new Formatter();
    private readonly imageReplacer = new ImageReplacer();
    private readonly chartReplacer = new ChartReplacer();
    private readonly hyperlinkReplacer = new HyperlinkReplacer();
    private readonly mediaReplacer = new MediaReplacer();

    public compile(
        file: File,
        prettifyXml?: string,
        overrides: readonly IXmlifyedFile[] = [],
    ): Zippable {
        const declaration = true;
        const indent = prettifyXml;
        const context: IContext = { fileData: file, stack: [] };

        const mapping: IXmlifyedFileMapping = {
            AppProperties: {
                data: xml(this.formatter.format(file.AppProperties, context), {
                    declaration,
                    indent,
                }),
                path: "docProps/app.xml",
            },
            Properties: {
                data: xml(this.formatter.format(file.CoreProperties, context), {
                    declaration,
                    indent,
                }),
                path: "docProps/core.xml",
            },
            FileRelationships: {
                data: xml(this.formatter.format(file.FileRelationships, context), {
                    declaration: false,
                }),
                path: "_rels/.rels",
            },
        };

        // Theme
        mapping["Theme"] = {
            data: xml(this.formatter.format(file.Theme, context), { declaration, indent }),
            path: "ppt/theme/theme1.xml",
        };

        // Table styles
        mapping["TableStyles"] = {
            data: xml(this.formatter.format(file.TableStyles, context), { declaration, indent }),
            path: "ppt/tableStyles.xml",
        };

        // Presentation properties
        mapping["PresProps"] = {
            data: xml(this.formatter.format(file.PresProps, context), { declaration, indent }),
            path: "ppt/presProps.xml",
        };

        // View properties
        mapping["ViewProps"] = {
            data: xml(this.formatter.format(file.ViewProps, context), { declaration, indent }),
            path: "ppt/viewProps.xml",
        };

        // Slide Master
        mapping["SlideMaster"] = {
            data: xml(this.formatter.format(file.SlideMaster, context), { declaration, indent }),
            path: "ppt/slideMasters/slideMaster1.xml",
        };

        // Slide Master relationships
        mapping["SlideMasterRelationships"] = {
            data: xml(this.formatter.format(file.SlideMasterRelationships, context), {
                declaration: false,
            }),
            path: "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        };

        // Slide Layout
        mapping["SlideLayout"] = {
            data: xml(this.formatter.format(file.SlideLayout, context), { declaration, indent }),
            path: "ppt/slideLayouts/slideLayout1.xml",
        };

        // Slide Layout relationships
        mapping["SlideLayoutRelationships"] = {
            data: xml(this.formatter.format(file.SlideLayoutRelationships, context), {
                declaration: false,
            }),
            path: "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
        };

        // Notes Master
        mapping["NotesMaster"] = {
            data: xml(this.formatter.format(new DefaultNotesMaster(), context), {
                declaration,
                indent,
            }),
            path: "ppt/notesMasters/notesMaster1.xml",
        };
        mapping["NotesMasterRelationships"] = {
            data: xml(this.formatter.format(file.NotesMasterRelationships, context), {
                declaration: false,
            }),
            path: "ppt/notesMasters/_rels/notesMaster1.xml.rels",
        };

        // Presentation + its relationships
        const presentationXml = xml(this.formatter.format(file.PresentationWrapper.View, context), {
            declaration,
            indent,
        });
        let currentImageCount = 0;

        const mediaData = this.imageReplacer.getMediaData(presentationXml, file.Media);
        mediaData.forEach((image) => {
            file.PresentationWrapper.Relationships.addRelationship(
                file.PresentationWrapper.Relationships.RelationshipCount + 1,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                `../media/${image.fileName}`,
            );
        });

        const replacedPresentationXml = this.imageReplacer.replace(
            presentationXml,
            mediaData,
            currentImageCount,
        );
        currentImageCount += mediaData.length;

        mapping["Presentation"] = {
            data: replacedPresentationXml,
            path: "ppt/presentation.xml",
        };

        mapping["PresentationRelationships"] = {
            data: xml(this.formatter.format(file.PresentationWrapper.Relationships, context), {
                declaration: false,
            }),
            path: "ppt/_rels/presentation.xml.rels",
        };

        // Slides — format BEFORE ContentTypes so ChartFrame.prepForXml() populates Charts
        for (let i = 0; i < file.Slides.length; i++) {
            const slideWrapper = file.SlideWrappers[i];
            const slideXml = xml(this.formatter.format(slideWrapper.View, context), {
                declaration,
                indent,
            });

            const slideMediaData = this.imageReplacer.getMediaData(slideXml, file.Media);
            slideMediaData.forEach((image) => {
                slideWrapper.Relationships.addRelationship(
                    slideWrapper.Relationships.RelationshipCount + 1,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    `../media/${image.fileName}`,
                );
            });

            let replacedSlideXml = this.imageReplacer.replace(
                slideXml,
                slideMediaData,
                currentImageCount,
            );
            currentImageCount += slideMediaData.length;

            // Chart placeholder replacement — only for charts present in this slide
            const chartPlaceholderRegex = /\{chart:([^}]+)\}/g;
            const slideChartKeys: string[] = [];
            let match: RegExpExecArray | null;
            const slideXmlForCharts = replacedSlideXml;
            while ((match = chartPlaceholderRegex.exec(slideXmlForCharts)) !== null) {
                if (!slideChartKeys.includes(match[1])) {
                    slideChartKeys.push(match[1]);
                }
            }

            if (slideChartKeys.length > 0) {
                const slideChartOffset = slideWrapper.Relationships.RelationshipCount + 1;
                const slideCharts = file.Charts.Array.filter((c) => slideChartKeys.includes(c.key));

                replacedSlideXml = this.chartReplacer.replace(
                    replacedSlideXml,
                    { Array: slideCharts } as unknown as ChartCollection,
                    slideChartOffset,
                );

                slideCharts.forEach((chartData, ci) => {
                    const globalIndex = file.Charts.Array.indexOf(chartData);
                    slideWrapper.Relationships.addRelationship(
                        slideChartOffset + ci,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                        `../charts/chart${globalIndex + 1}.xml`,
                    );
                });
            }

            // Hyperlink placeholder replacement
            const hlinkPlaceholderRegex = /\{hlink:([^}]+)\}/g;
            const slideHlinkKeys: string[] = [];
            let hlinkMatch: RegExpExecArray | null;
            while ((hlinkMatch = hlinkPlaceholderRegex.exec(replacedSlideXml)) !== null) {
                if (!slideHlinkKeys.includes(hlinkMatch[1])) {
                    slideHlinkKeys.push(hlinkMatch[1]);
                }
            }

            if (slideHlinkKeys.length > 0) {
                const slideHlinks = file.Hyperlinks.Array.filter((h) => slideHlinkKeys.includes(h.key));
                const hlinkOffset = slideWrapper.Relationships.RelationshipCount + 1;

                replacedSlideXml = this.hyperlinkReplacer.replace(
                    replacedSlideXml,
                    slideHlinks,
                    hlinkOffset,
                );

                slideHlinks.forEach((hlink, hi) => {
                    slideWrapper.Relationships.addRelationship(
                        hlinkOffset + hi,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                        hlink.url,
                        "External",
                    );
                });
            }

            // Media (video/audio) placeholder replacement
            const slideMediaRefs = this.mediaReplacer.getMediaRefs(replacedSlideXml, file.Media);
            if (slideMediaRefs.length > 0) {
                const mediaOffset = slideWrapper.Relationships.RelationshipCount + 1;

                replacedSlideXml = this.mediaReplacer.replace(
                    replacedSlideXml,
                    slideMediaRefs,
                    mediaOffset,
                );

                slideMediaRefs.forEach((media, mi) => {
                    slideWrapper.Relationships.addRelationship(
                        mediaOffset + mi,
                        "http://schemas.microsoft.com/office/powerpoint/2010/relationships/media",
                        `../media/${media.fileName}`,
                    );
                });
            }

            mapping[`Slide${i}`] = {
                data: replacedSlideXml,
                path: `ppt/slides/slide${i + 1}.xml`,
            };

            mapping[`SlideRelationships${i}`] = {
                data: xml(this.formatter.format(slideWrapper.Relationships, context), {
                    declaration: false,
                }),
                path: `ppt/slides/_rels/slide${i + 1}.xml.rels`,
            };
        }

        // ContentTypes — AFTER slides so Charts.Array is populated
        file.Charts.Array.forEach((_, i) => {
            file.ContentTypes.addChart(i + 1);
        });
        mapping["ContentTypes"] = {
            data: xml(this.formatter.format(file.ContentTypes, context), {
                declaration: false,
            }),
            path: "[Content_Types].xml",
        };

        // Convert mapping to Zippable
        const files: Zippable = {};
        for (const key of Object.keys(mapping)) {
            const entry = mapping[key];
            files[entry.path] = textToUint8Array(entry.data);
        }

        // Add overrides
        for (const override of overrides) {
            files[override.path] =
                override.data instanceof Uint8Array
                    ? override.data
                    : textToUint8Array(override.data);
        }

        // Add chart parts
        for (let i = 0; i < file.Charts.Array.length; i++) {
            const chartData = file.Charts.Array[i];
            files[`ppt/charts/chart${i + 1}.xml`] = textToUint8Array(
                xml(this.formatter.format(chartData.chartSpace, context), {
                    declaration,
                    indent,
                }),
            );
            files[`ppt/charts/_rels/chart${i + 1}.xml.rels`] = textToUint8Array(
                xml(
                    {
                        Relationships: {
                            _attr: {
                                xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
                            },
                        },
                    },
                    { declaration: { encoding: "UTF-8", standalone: "yes" } },
                ),
            );
        }

        // Add notes slides
        for (let i = 0; i < file.NotesSlides.length; i++) {
            const notesSlide = file.NotesSlides[i];
            files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = textToUint8Array(
                xml(this.formatter.format(notesSlide, context), {
                    declaration,
                    indent,
                }),
            );
            files[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = textToUint8Array(
                xml(
                    {
                        Relationships: {
                            _attr: {
                                xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
                            },
                        },
                    },
                    { declaration: { encoding: "UTF-8", standalone: "yes" } },
                ),
            );
        }

        // Add media files (STORE compression)
        for (const image of file.Media.Array) {
            files[`ppt/media/${image.fileName}`] = [image.data, { level: 0 }];
            if (image.type === "svg" && "fallback" in image) {
                const fallback = (
                    image as import("@file/media/data").IMediaData & {
                        readonly fallback: { readonly fileName: string; readonly data: Uint8Array };
                    }
                ).fallback;
                files[`ppt/media/${fallback.fileName}`] = [fallback.data, { level: 0 }];
            }
        }

        return files;
    }
}

function textToUint8Array(data: string | Uint8Array): Uint8Array {
    if (data instanceof Uint8Array) {
        return data;
    }
    return new TextEncoder().encode(data);
}
