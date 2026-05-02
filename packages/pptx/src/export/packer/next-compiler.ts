import { Formatter } from "@export/formatter";
import type { File } from "@file/file";
import type { IContext } from "@file/xml-components";
import { xml } from "@office-open/xml";
import type { Zippable } from "fflate";

import { ImageReplacer } from "./image-replacer";

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
            ContentTypes: {
                data: xml(this.formatter.format(file.ContentTypes, context), {
                    declaration: false,
                }),
                path: "[Content_Types].xml",
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

        // Slides
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

            const replacedSlideXml = this.imageReplacer.replace(
                slideXml,
                slideMediaData,
                currentImageCount,
            );
            currentImageCount += slideMediaData.length;

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
