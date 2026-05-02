import type { File } from "@file/file";
import { obfuscate } from "@file/fonts/obfuscate-ttf-to-odttf";
import {
    DEFAULT_COLORS_XML,
    DEFAULT_LAYOUT_XML,
    DEFAULT_STYLE_XML,
} from "@file/smartart/built-in-definitions";
import { xml } from "@office-open/xml";
/**
 * Compiler module for converting File objects into OOXML ZIP archives.
 *
 * @module
 */
import type { Zippable } from "fflate";
import { textToUint8Array, toUint8Array } from "undio";

import { Formatter } from "../formatter";
import { ChartReplacer } from "./chart-replacer";
import { ImageReplacer } from "./image-replacer";
import { NumberingReplacer } from "./numbering-replacer";
import type { PrettifyType } from "./packer";
import { SmartArtReplacer } from "./smartart-replacer";

/**
 * Represents a serialized XML file with its path in the OOXML package.
 *
 * @property data - The XML content as a string
 * @property path - The file path within the ZIP archive (e.g., "word/document.xml")
 */
export interface IXmlifyedFile {
    readonly data: string | Uint8Array;
    readonly path: string;
}

/**
 * Complete mapping of all XML files in an OOXML document package.
 *
 * This type represents the full structure of a .docx file, including the main
 * document, styles, relationships, headers, footers, and metadata files.
 */
interface IXmlifyedFileMapping {
    /** Main document content (word/document.xml) */
    readonly Document: IXmlifyedFile;
    /** Style definitions (word/styles.xml) */
    readonly Styles: IXmlifyedFile;
    /** Core document properties (docProps/core.xml) */
    readonly Properties: IXmlifyedFile;
    /** Numbering definitions (word/numbering.xml) */
    readonly Numbering: IXmlifyedFile;
    /** Document relationships (word/_rels/document.xml.rels) */
    readonly Relationships: IXmlifyedFile;
    /** Package-level relationships (_rels/.rels) */
    readonly FileRelationships: IXmlifyedFile;
    /** Header content files */
    readonly Headers: readonly IXmlifyedFile[];
    /** Footer content files */
    readonly Footers: readonly IXmlifyedFile[];
    /** Header relationship files */
    readonly HeaderRelationships: readonly IXmlifyedFile[];
    /** Footer relationship files */
    readonly FooterRelationships: readonly IXmlifyedFile[];
    /** Content types mapping ([Content_Types].xml) */
    readonly ContentTypes: IXmlifyedFile;
    /** Custom document properties (docProps/custom.xml) */
    readonly CustomProperties: IXmlifyedFile;
    /** Application properties (docProps/app.xml) */
    readonly AppProperties: IXmlifyedFile;
    /** Footnotes content (word/footnotes.xml) */
    readonly FootNotes: IXmlifyedFile;
    /** Footnotes relationships (word/_rels/footnotes.xml.rels) */
    readonly FootNotesRelationships: IXmlifyedFile;
    /** Endnotes content (word/endnotes.xml) */
    readonly Endnotes: IXmlifyedFile;
    /** Endnotes relationships (word/_rels/endnotes.xml.rels) */
    readonly EndnotesRelationships: IXmlifyedFile;
    /** Document settings (word/settings.xml) */
    readonly Settings: IXmlifyedFile;
    /** Comments content (word/comments.xml) */
    readonly Comments?: IXmlifyedFile;
    /** Comments relationships (word/_rels/comments.xml.rels) */
    readonly CommentsRelationships?: IXmlifyedFile;
    /** Font table (word/fontTable.xml) */
    readonly FontTable?: IXmlifyedFile;
    /** Font table relationships (word/_rels/fontTable.xml.rels) */
    readonly FontTableRelationships?: IXmlifyedFile;
    /** Bibliography content (word/bibliography.xml) */
    readonly Bibliography?: IXmlifyedFile;
    /** Chart XML parts (word/charts/chart{n}.xml) */
    readonly Charts?: readonly IXmlifyedFile[];
    /** Diagram data XML parts (word/diagrams/data{n}.xml) */
    readonly DiagramData?: readonly IXmlifyedFile[];
    /** Diagram layout XML parts (word/diagrams/layout{n}.xml) */
    readonly DiagramLayout?: readonly IXmlifyedFile[];
    /** Diagram style XML parts (word/diagrams/quickStyle{n}.xml) */
    readonly DiagramStyle?: readonly IXmlifyedFile[];
    /** Diagram colors XML parts (word/diagrams/colors{n}.xml) */
    readonly DiagramColors?: readonly IXmlifyedFile[];
    /** AltChunk parts (word/afchunks/afchunk{n}.{ext}) */
    readonly AltChunks?: readonly IXmlifyedFile[];
    /** SubDoc parts (word/subdocs/subdoc{n}.docx) */
    readonly SubDocs?: readonly IXmlifyedFile[];
}

/**
 * Compiles File objects into OOXML-compliant ZIP file data.
 *
 * The Compiler is responsible for converting the internal document representation
 * into the complete set of XML files required for a .docx document, managing
 * relationships, images, fonts, and all other document components.
 *
 * @example
 * ```typescript
 * const compiler = new Compiler();
 * const files = compiler.compile(file, PrettifyType.WITH_2_BLANKS);
 * ```
 */

export class Compiler {
    private readonly formatter: Formatter;
    private readonly imageReplacer: ImageReplacer;
    private readonly numberingReplacer: NumberingReplacer;
    private readonly chartReplacer: ChartReplacer;
    private readonly smartArtReplacer: SmartArtReplacer;

    /**
     * Creates a new Compiler instance.
     *
     * Initializes the formatter and replacer utilities used during compilation.
     */
    public constructor() {
        this.formatter = new Formatter();
        this.imageReplacer = new ImageReplacer();
        this.numberingReplacer = new NumberingReplacer();
        this.chartReplacer = new ChartReplacer();
        this.smartArtReplacer = new SmartArtReplacer();
    }

    /**
     * Compiles a File object into a flat file map suitable for fflate zipSync.
     *
     * This method orchestrates the entire compilation process:
     * - Converts all document components to XML
     * - Manages image and numbering placeholder replacements
     * - Creates relationship files
     * - Packages fonts and media files
     *
     * Media files (images, SVGs) use level 0 (STORE) to avoid redundant
     * compression on already-compressed formats. XML files use the default
     * DEFLATE compression set by the caller via zipSync options.
     *
     * @param file - The document to compile
     * @param prettifyXml - Optional XML formatting style
     * @param overrides - Optional custom XML file overrides
     * @returns A Zippable object mapping file paths to their content
     */
    public compile(
        file: File,
        prettifyXml?: (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Zippable {
        const files: Zippable = {};

        // Cache format() results to avoid duplicate formatting of Header/Footer views.
        // HeaderRelationships/FooterRelationships runs before Headers/Footers in the map,
        // So we cache the formatted View data there and reuse it later.
        const headerFormattedViews = new Map<number, string>();
        const footerFormattedViews = new Map<number, string>();

        const xmlifiedFileMapping = this.xmlifyFile(
            file,
            headerFormattedViews,
            footerFormattedViews,
            prettifyXml,
        );
        const map = new Map<string, IXmlifyedFile | readonly IXmlifyedFile[]>(
            Object.entries(xmlifiedFileMapping),
        );

        for (const [, obj] of map) {
            if (Array.isArray(obj)) {
                for (const subFile of obj as readonly IXmlifyedFile[]) {
                    files[subFile.path] =
                        typeof subFile.data === "string"
                            ? textToUint8Array(subFile.data)
                            : subFile.data;
                }
            } else {
                const fileObj = obj as IXmlifyedFile;
                files[fileObj.path] =
                    typeof fileObj.data === "string"
                        ? textToUint8Array(fileObj.data)
                        : fileObj.data;
            }
        }

        for (const subFile of overrides) {
            files[subFile.path] =
                typeof subFile.data === "string" ? textToUint8Array(subFile.data) : subFile.data;
        }

        // Media files: use STORE (level 0) for already-compressed formats
        for (const mediaData of file.Media.Array) {
            files[`word/media/${mediaData.fileName}`] = [
                toUint8Array(mediaData.data),
                { level: 0 },
            ];
            if (mediaData.type === "svg") {
                files[`word/media/${mediaData.fallback.fileName}`] = [
                    toUint8Array(mediaData.fallback.data),
                    { level: 0 },
                ];
            }
        }

        // Font files: use higher compression for binary font data
        for (const { data: buffer, name, fontKey } of file.FontTable.fontOptionsWithKey) {
            const [nameWithoutExtension] = name.split(".");
            files[`word/fonts/${nameWithoutExtension}.odttf`] = obfuscate(buffer, fontKey);
        }

        return files;
    }

    private xmlifyFile(
        file: File,
        headerFormattedViews: Map<number, string>,
        footerFormattedViews: Map<number, string>,
        prettify?: (typeof PrettifyType)[keyof typeof PrettifyType],
    ): IXmlifyedFileMapping {
        const documentRelationshipCount = file.Document.Relationships.RelationshipCount + 1;

        const documentXmlData = xml(
            this.formatter.format(file.Document.View, {
                fileData: file,
                file,
                stack: [],
                viewWrapper: file.Document,
            }),
            {
                declaration: {
                    encoding: "UTF-8",
                    standalone: "yes",
                },
                indent: prettify,
            },
        );

        const commentRelationshipCount = file.Comments.Relationships.RelationshipCount + 1;
        const commentXmlData = xml(
            this.formatter.format(file.Comments, {
                fileData: file,
                file,
                stack: [],
                viewWrapper: {
                    Relationships: file.Comments.Relationships,
                    View: file.Comments,
                },
            }),
            {
                declaration: {
                    encoding: "UTF-8",
                    standalone: "yes",
                },
                indent: prettify,
            },
        );

        const footnoteRelationshipCount = file.FootNotes.Relationships.RelationshipCount + 1;
        const footnoteXmlData = xml(
            this.formatter.format(file.FootNotes.View, {
                fileData: file,
                file,
                stack: [],
                viewWrapper: file.FootNotes,
            }),
            {
                declaration: {
                    encoding: "UTF-8",
                    standalone: "yes",
                },
                indent: prettify,
            },
        );

        const documentMediaDatas = this.imageReplacer.getMediaData(documentXmlData, file.Media);
        const commentMediaDatas = this.imageReplacer.getMediaData(commentXmlData, file.Media);
        const footnoteMediaDatas = this.imageReplacer.getMediaData(footnoteXmlData, file.Media);

        return {
            AppProperties: {
                data: xml(
                    this.formatter.format(file.AppProperties, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Document,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                            standalone: "yes",
                        },
                        indent: prettify,
                    },
                ),
                path: "docProps/app.xml",
            },
            Comments: {
                data: (() => {
                    const xmlData = this.imageReplacer.replace(
                        commentXmlData,
                        commentMediaDatas,
                        commentRelationshipCount,
                    );
                    const referenedXmlData = this.numberingReplacer.replace(
                        xmlData,
                        file.Numbering.ConcreteNumbering,
                    );
                    return referenedXmlData;
                })(),
                path: "word/comments.xml",
            },
            CommentsRelationships: {
                data: (() => {
                    commentMediaDatas.forEach((mediaData, i) => {
                        file.Comments.Relationships.addRelationship(
                            commentRelationshipCount + i,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            `media/${mediaData.fileName}`,
                        );
                    });
                    return xml(
                        this.formatter.format(file.Comments.Relationships, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: {
                                Relationships: file.Comments.Relationships,
                                View: file.Comments,
                            },
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                            },
                            indent: prettify,
                        },
                    );
                })(),
                path: "word/_rels/comments.xml.rels",
            },
            ContentTypes: {
                data: (() => {
                    // Register chart and diagram content types BEFORE serialization
                    file.Charts.Array.forEach((_, i) => {
                        file.ContentTypes.addChart(i + 1);
                    });
                    file.SmartArts.Array.forEach((_, i) => {
                        file.ContentTypes.addDiagramData(i + 1);
                        file.ContentTypes.addDiagramLayout(i + 1);
                        file.ContentTypes.addDiagramStyle(i + 1);
                        file.ContentTypes.addDiagramColors(i + 1);
                    });
                    return xml(
                        this.formatter.format(file.ContentTypes, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: file.Document,
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                            },
                            indent: prettify,
                        },
                    );
                })(),
                path: "[Content_Types].xml",
            },
            CustomProperties: {
                data: xml(
                    this.formatter.format(file.CustomProperties, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Document,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                            standalone: "yes",
                        },
                        indent: prettify,
                    },
                ),
                path: "docProps/custom.xml",
            },
            Document: {
                data: (() => {
                    let xmlData = this.imageReplacer.replace(
                        documentXmlData,
                        documentMediaDatas,
                        documentRelationshipCount,
                    );
                    xmlData = this.chartReplacer.replace(
                        xmlData,
                        file.Charts,
                        documentRelationshipCount,
                    );
                    const smartArtDataOffset =
                        documentRelationshipCount +
                        documentMediaDatas.length +
                        file.Charts.Array.length;
                    xmlData = this.smartArtReplacer.replace(
                        xmlData,
                        file.SmartArts,
                        smartArtDataOffset,
                    );
                    const referenedXmlData = this.numberingReplacer.replace(
                        xmlData,
                        file.Numbering.ConcreteNumbering,
                    );
                    return referenedXmlData;
                })(),
                path: "word/document.xml",
            },
            Endnotes: {
                data: xml(
                    this.formatter.format(file.Endnotes.View, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Endnotes,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                        },
                        indent: prettify,
                    },
                ),
                path: "word/endnotes.xml",
            },
            EndnotesRelationships: {
                data: xml(
                    this.formatter.format(file.Endnotes.Relationships, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Endnotes,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                        },
                        indent: prettify,
                    },
                ),
                path: "word/_rels/endnotes.xml.rels",
            },
            FileRelationships: {
                data: xml(
                    this.formatter.format(file.FileRelationships, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Document,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                        },
                        indent: prettify,
                    },
                ),
                path: "_rels/.rels",
            },
            FontTable: {
                data: xml(
                    this.formatter.format(file.FontTable.View, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Document,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                            standalone: "yes",
                        },
                        indent: prettify,
                    },
                ),
                path: "word/fontTable.xml",
            },
            FontTableRelationships: {
                data: (() =>
                    xml(
                        this.formatter.format(file.FontTable.Relationships, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: file.Document,
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                            },
                            indent: prettify,
                        },
                    ))(),
                path: "word/_rels/fontTable.xml.rels",
            },
            FootNotes: {
                data: (() => {
                    const xmlData = this.imageReplacer.replace(
                        footnoteXmlData,
                        footnoteMediaDatas,
                        footnoteRelationshipCount,
                    );
                    const referenedXmlData = this.numberingReplacer.replace(
                        xmlData,
                        file.Numbering.ConcreteNumbering,
                    );
                    return referenedXmlData;
                })(),
                path: "word/footnotes.xml",
            },
            FootNotesRelationships: {
                data: (() => {
                    footnoteMediaDatas.forEach((mediaData, i) => {
                        file.FootNotes.Relationships.addRelationship(
                            footnoteRelationshipCount + i,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            `media/${mediaData.fileName}`,
                        );
                    });
                    return xml(
                        this.formatter.format(file.FootNotes.Relationships, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: file.FootNotes,
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                            },
                            indent: prettify,
                        },
                    );
                })(),
                path: "word/_rels/footnotes.xml.rels",
            },
            FooterRelationships: file.Footers.map((footerWrapper, index) => {
                const formatted = this.formatter.format(footerWrapper.View, {
                    fileData: file,
                    file,
                    stack: [],
                    viewWrapper: footerWrapper,
                });
                const xmlData = xml(formatted, {
                    declaration: {
                        encoding: "UTF-8",
                    },
                    indent: prettify,
                });
                // Cache for reuse in Footers section
                footerFormattedViews.set(index, xmlData);
                const mediaDatas = this.imageReplacer.getMediaData(xmlData, file.Media);

                mediaDatas.forEach((mediaData, i) => {
                    footerWrapper.Relationships.addRelationship(
                        i,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        `media/${mediaData.fileName}`,
                    );
                });

                return {
                    data: xml(
                        this.formatter.format(footerWrapper.Relationships, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: footerWrapper,
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                            },
                            indent: prettify,
                        },
                    ),
                    path: `word/_rels/footer${index + 1}.xml.rels`,
                };
            }),
            Footers: file.Footers.map((_footerWrapper, index) => {
                const tempXmlData = footerFormattedViews.get(index)!;
                const mediaDatas = this.imageReplacer.getMediaData(tempXmlData, file.Media);
                // TODO: 0 needs to be changed when headers get relationships of their own
                const xmlData = this.imageReplacer.replace(tempXmlData, mediaDatas, 0);

                const referenedXmlData = this.numberingReplacer.replace(
                    xmlData,
                    file.Numbering.ConcreteNumbering,
                );

                return {
                    data: referenedXmlData,
                    path: `word/footer${index + 1}.xml`,
                };
            }),
            HeaderRelationships: file.Headers.map((headerWrapper, index) => {
                const formatted = this.formatter.format(headerWrapper.View, {
                    fileData: file,
                    file,
                    stack: [],
                    viewWrapper: headerWrapper,
                });
                const xmlData = xml(formatted, {
                    declaration: {
                        encoding: "UTF-8",
                    },
                    indent: prettify,
                });
                // Cache for reuse in Headers section
                headerFormattedViews.set(index, xmlData);
                const mediaDatas = this.imageReplacer.getMediaData(xmlData, file.Media);

                mediaDatas.forEach((mediaData, i) => {
                    headerWrapper.Relationships.addRelationship(
                        i,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        `media/${mediaData.fileName}`,
                    );
                });

                return {
                    data: xml(
                        this.formatter.format(headerWrapper.Relationships, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: headerWrapper,
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                            },
                            indent: prettify,
                        },
                    ),
                    path: `word/_rels/header${index + 1}.xml.rels`,
                };
            }),
            Headers: file.Headers.map((_headerWrapper, index) => {
                const tempXmlData = headerFormattedViews.get(index)!;
                const mediaDatas = this.imageReplacer.getMediaData(tempXmlData, file.Media);
                // TODO: 0 needs to be changed when headers get relationships of their own
                const xmlData = this.imageReplacer.replace(tempXmlData, mediaDatas, 0);

                const referenedXmlData = this.numberingReplacer.replace(
                    xmlData,
                    file.Numbering.ConcreteNumbering,
                );

                return {
                    data: referenedXmlData,
                    path: `word/header${index + 1}.xml`,
                };
            }),
            Numbering: {
                data: xml(
                    this.formatter.format(file.Numbering, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Document,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                            standalone: "yes",
                        },
                        indent: prettify,
                    },
                ),
                path: "word/numbering.xml",
            },
            Properties: {
                data: xml(
                    this.formatter.format(file.CoreProperties, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Document,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                            standalone: "yes",
                        },
                        indent: prettify,
                    },
                ),
                path: "docProps/core.xml",
            },
            Relationships: {
                data: (() => {
                    documentMediaDatas.forEach((mediaData, i) => {
                        file.Document.Relationships.addRelationship(
                            documentRelationshipCount + i,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            `media/${mediaData.fileName}`,
                        );
                    });

                    // Chart relationships
                    const chartOffset = documentRelationshipCount + documentMediaDatas.length;
                    file.Charts.Array.forEach((_chartData, i) => {
                        file.Document.Relationships.addRelationship(
                            chartOffset + i,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                            `charts/chart${i + 1}.xml`,
                        );
                    });

                    // SmartArt relationships (data + layout/style/color internal)
                    this.smartArtReplacer.addRelationships(
                        file.SmartArts,
                        (id, type, target, targetMode) => {
                            file.Document.Relationships.addRelationship(
                                id,
                                type,
                                target,
                                targetMode as any,
                            );
                        },
                        documentRelationshipCount,
                        documentMediaDatas.length + file.Charts.Array.length,
                    );

                    file.Document.Relationships.addRelationship(
                        file.Document.Relationships.RelationshipCount + 1,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
                        "fontTable.xml",
                    );

                    return xml(
                        this.formatter.format(file.Document.Relationships, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: file.Document,
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                            },
                            indent: prettify,
                        },
                    );
                })(),
                path: "word/_rels/document.xml.rels",
            },
            Settings: {
                data: xml(
                    this.formatter.format(file.Settings, {
                        fileData: file,
                        file,
                        stack: [],
                        viewWrapper: file.Document,
                    }),
                    {
                        declaration: {
                            encoding: "UTF-8",
                            standalone: "yes",
                        },
                        indent: prettify,
                    },
                ),
                path: "word/settings.xml",
            },
            Styles: {
                data: (() => {
                    const xmlStyles = xml(
                        this.formatter.format(file.Styles, {
                            fileData: file,
                            file,
                            stack: [],
                            viewWrapper: file.Document,
                        }),
                        {
                            declaration: {
                                encoding: "UTF-8",
                                standalone: "yes",
                            },
                            indent: prettify,
                        },
                    );
                    const referencedXmlStyles = this.numberingReplacer.replace(
                        xmlStyles,
                        file.Numbering.ConcreteNumbering,
                    );
                    return referencedXmlStyles;
                })(),
                path: "word/styles.xml",
            },
            ...(file.Bibliography
                ? {
                      Bibliography: {
                          data: xml(
                              this.formatter.format(file.Bibliography, {
                                  fileData: file,
                                  file,
                                  stack: [],
                                  viewWrapper: {
                                      Relationships: file.Bibliography.Relationships,
                                      View: file.Bibliography,
                                  },
                              }),
                              {
                                  declaration: {
                                      encoding: "UTF-8",
                                      standalone: "yes",
                                  },
                                  indent: prettify,
                              },
                          ),
                          path: "word/bibliography.xml",
                      },
                  }
                : {}),
            ...(file.Charts.Array.length > 0
                ? {
                      Charts: file.Charts.Array.flatMap((chartData, i) => [
                          {
                              data: xml(
                                  this.formatter.format(chartData.chartSpace, {
                                      fileData: file,
                                      file,
                                      stack: [],
                                      viewWrapper: file.Document,
                                  }),
                                  {
                                      declaration: {
                                          encoding: "UTF-8",
                                          standalone: "yes",
                                      },
                                      indent: prettify,
                                  },
                              ),
                              path: `word/charts/chart${i + 1}.xml`,
                          },
                          {
                              data: xml(
                                  {
                                      Relationships: {
                                          _attr: {
                                              xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
                                          },
                                      },
                                  },
                                  { declaration: { encoding: "UTF-8", standalone: "yes" } },
                              ),
                              path: `word/charts/_rels/chart${i + 1}.xml.rels`,
                          },
                      ]),
                  }
                : {}),
            ...(file.SmartArts.Array.length > 0
                ? {
                      DiagramData: file.SmartArts.Array.map((smartArtData, i) => ({
                          data: xml(
                              this.formatter.format(smartArtData.dataModel, {
                                  fileData: file,
                                  file,
                                  stack: [],
                                  viewWrapper: file.Document,
                              }),
                              {
                                  declaration: {
                                      encoding: "UTF-8",
                                      standalone: "yes",
                                  },
                                  indent: prettify,
                              },
                          ),
                          path: `word/diagrams/data${i + 1}.xml`,
                      })),
                      DiagramLayout: file.SmartArts.Array.map((_smartArtData, i) => ({
                          data: DEFAULT_LAYOUT_XML,
                          path: `word/diagrams/layout${i + 1}.xml`,
                      })),
                      DiagramStyle: file.SmartArts.Array.map((_smartArtData, i) => ({
                          data: DEFAULT_STYLE_XML,
                          path: `word/diagrams/quickStyle${i + 1}.xml`,
                      })),
                      DiagramColors: file.SmartArts.Array.map((_smartArtData, i) => ({
                          data: DEFAULT_COLORS_XML,
                          path: `word/diagrams/colors${i + 1}.xml`,
                      })),
                  }
                : {}),
            ...(file.AltChunks.Array.length > 0
                ? {
                      AltChunks: file.AltChunks.Array.map((altChunkData) => ({
                          data: altChunkData.data,
                          path: `word/${altChunkData.path}`,
                      })),
                  }
                : {}),
            ...(file.SubDocs.Array.length > 0
                ? {
                      SubDocs: file.SubDocs.Array.map((subDocData) => ({
                          data: subDocData.data,
                          path: `word/${subDocData.path}`,
                      })),
                  }
                : {}),
        };
    }
}
