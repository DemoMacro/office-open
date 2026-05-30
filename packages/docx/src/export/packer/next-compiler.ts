import type { File } from "@file/file";
import { obfuscate } from "@file/fonts/obfuscate-ttf-to-odttf";
import {
  DEFAULT_DRAWING_XML,
  getColorXml,
  getLayoutXml,
  getStyleXml,
} from "@file/smartart/built-in-definitions";
import {
  addSmartArtRelationships,
  getReferencedMedia,
  hasPlaceholders,
  replaceChartPlaceholders,
  replaceImagePlaceholders,
  replaceSmartArtPlaceholders,
} from "@office-open/core";
import type { Context, XmlifyedFile, Zippable } from "@office-open/core";
import { ZIP_STORED_LEVEL } from "@office-open/core";
import { xml } from "@office-open/xml";
import { textToUint8Array, toUint8Array } from "undio";

import { Formatter } from "../formatter";
import { replaceNumberingPlaceholders } from "./numbering-placeholders";

/** Extended context with docx-specific properties passed through the XML tree. */
type DocxContext = Context & {
  file: File;
  viewWrapper?: any;
  headerFormattedViews?: Map<number, string>;
  footerFormattedViews?: Map<number, string>;
};

/**
 * Complete mapping of all XML files in an OOXML document package.
 *
 * This type represents the full structure of a .docx file, including the main
 * document, styles, relationships, headers, footers, and metadata files.
 */
interface XmlifyedFileMapping {
  /** Main document content (word/document.xml) */
  readonly Document: XmlifyedFile;
  /** Style definitions (word/styles.xml) */
  readonly Styles: XmlifyedFile;
  /** Core document properties (docProps/core.xml) */
  readonly Properties: XmlifyedFile;
  /** Numbering definitions (word/numbering.xml) */
  readonly Numbering: XmlifyedFile;
  /** Document relationships (word/_rels/document.xml.rels) */
  readonly Relationships: XmlifyedFile;
  /** Package-level relationships (_rels/.rels) */
  readonly FileRelationships: XmlifyedFile;
  /** Header content files */
  readonly Headers: readonly XmlifyedFile[];
  /** Footer content files */
  readonly Footers: readonly XmlifyedFile[];
  /** Header relationship files */
  readonly HeaderRelationships: readonly XmlifyedFile[];
  /** Footer relationship files */
  readonly FooterRelationships: readonly XmlifyedFile[];
  /** Content types mapping ([Content_Types].xml) */
  readonly ContentTypes: XmlifyedFile;
  /** Custom document properties (docProps/custom.xml) */
  readonly CustomProperties: XmlifyedFile;
  /** Application properties (docProps/app.xml) */
  readonly AppProperties: XmlifyedFile;
  /** Footnotes content (word/footnotes.xml) */
  readonly FootNotes: XmlifyedFile;
  /** Footnotes relationships (word/_rels/footnotes.xml.rels) */
  readonly FootNotesRelationships: XmlifyedFile;
  /** Endnotes content (word/endnotes.xml) */
  readonly Endnotes: XmlifyedFile;
  /** Endnotes relationships (word/_rels/endnotes.xml.rels) */
  readonly EndnotesRelationships: XmlifyedFile;
  /** Document settings (word/settings.xml) */
  readonly Settings: XmlifyedFile;
  /** Comments content (word/comments.xml) */
  readonly Comments?: XmlifyedFile;
  /** Comments relationships (word/_rels/comments.xml.rels) */
  readonly CommentsRelationships?: XmlifyedFile;
  /** Font table (word/fontTable.xml) */
  readonly FontTable?: XmlifyedFile;
  /** Font table relationships (word/_rels/fontTable.xml.rels) */
  readonly FontTableRelationships?: XmlifyedFile;
  /** Bibliography content (word/bibliography.xml) */
  readonly Bibliography?: XmlifyedFile;
  /** Chart XML parts (word/charts/chart{n}.xml) */
  readonly Charts?: readonly XmlifyedFile[];
  /** Diagram data XML parts (word/diagrams/data{n}.xml) */
  readonly DiagramData?: readonly XmlifyedFile[];
  /** Diagram layout XML parts (word/diagrams/layout{n}.xml) */
  readonly DiagramLayout?: readonly XmlifyedFile[];
  /** Diagram style XML parts (word/diagrams/quickStyle{n}.xml) */
  readonly DiagramStyle?: readonly XmlifyedFile[];
  /** Diagram colors XML parts (word/diagrams/colors{n}.xml) */
  readonly DiagramColors?: readonly XmlifyedFile[];
  /** Diagram drawing XML parts (word/diagrams/drawing{n}.xml) */
  readonly DiagramDrawing?: readonly XmlifyedFile[];
  /** AltChunk parts (word/afchunks/afchunk{n}.{ext}) */
  readonly AltChunks?: readonly XmlifyedFile[];
  /** SubDoc parts (word/subdocs/subdoc{n}.docx) */
  readonly SubDocs?: readonly XmlifyedFile[];
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
 * const files = compiler.compile(file);
 * ```
 */

export class Compiler {
  private readonly formatter: Formatter;

  public constructor() {
    this.formatter = new Formatter();
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
   * @param overrides - Optional custom XML file overrides
   * @returns A Zippable object mapping file paths to their content
   */
  public compile(file: File, overrides: readonly XmlifyedFile[] = []): Zippable {
    const files: Zippable = {};

    // Cache format() results to avoid duplicate formatting of Header/Footer views.
    // HeaderRelationships/FooterRelationships runs before Headers/Footers in the map,
    // So we cache the formatted View data there and reuse it later.
    const headerFormattedViews = new Map<number, string>();
    const footerFormattedViews = new Map<number, string>();

    const xmlifiedFileMapping = this.xmlifyFile(file, headerFormattedViews, footerFormattedViews);
    const map = new Map<string, XmlifyedFile | readonly XmlifyedFile[]>(
      Object.entries(xmlifiedFileMapping),
    );

    for (const [, obj] of map) {
      if (Array.isArray(obj)) {
        for (const subFile of obj as readonly XmlifyedFile[]) {
          files[subFile.path] =
            typeof subFile.data === "string" ? textToUint8Array(subFile.data) : subFile.data;
        }
      } else {
        const fileObj = obj as XmlifyedFile;
        files[fileObj.path] =
          typeof fileObj.data === "string" ? textToUint8Array(fileObj.data) : fileObj.data;
      }
    }

    for (const subFile of overrides) {
      files[subFile.path] =
        typeof subFile.data === "string" ? textToUint8Array(subFile.data) : subFile.data;
    }

    // Media files: use STORE for already-compressed formats (JPEG, PNG, GIF, etc.)
    for (const mediaData of file.media.array) {
      files[`word/media/${mediaData.fileName}`] = [
        toUint8Array(mediaData.data),
        { level: ZIP_STORED_LEVEL },
      ];
      if (mediaData.type === "svg") {
        files[`word/media/${mediaData.fallback.fileName}`] = [
          toUint8Array(mediaData.fallback.data),
          { level: ZIP_STORED_LEVEL },
        ];
      }
    }

    // Font files: use higher compression for binary font data
    for (const { data: buffer, name, fontKey } of file.fontTable.fontOptionsWithKey) {
      const [nameWithoutExtension] = name.split(".");
      files[`word/fonts/${nameWithoutExtension}.odttf`] = obfuscate(buffer, fontKey);
    }

    return files;
  }

  private xmlifyFile(
    file: File,
    headerFormattedViews: Map<number, string>,
    footerFormattedViews: Map<number, string>,
  ): XmlifyedFileMapping {
    // Factory to build typed context objects with docx-specific extensions.
    const mkCtx = (viewWrapper?: DocxContext["viewWrapper"]): DocxContext => ({
      fileData: file,
      file,
      stack: [],
      viewWrapper,
    });

    const documentRelationshipCount = file.document.relationships.relationshipCount + 1;

    const documentXmlData = this.formatter.formatToXml(file.document.view, mkCtx(file.document), {
      encoding: "UTF-8",
      standalone: "yes",
    });

    const commentRelationshipCount = file.comments.relationships.relationshipCount + 1;
    const commentXmlData = this.formatter.formatToXml(
      file.comments,
      mkCtx({ relationships: file.comments.relationships, view: file.comments }),
      { encoding: "UTF-8", standalone: "yes" },
    );

    const footnoteRelationshipCount = file.footNotes.relationships.relationshipCount + 1;
    const footnoteXmlData = this.formatter.formatToXml(file.footNotes.view, mkCtx(file.footNotes), {
      encoding: "UTF-8",
      standalone: "yes",
    });

    const documentMediaDatas = hasPlaceholders(documentXmlData)
      ? getReferencedMedia(documentXmlData, file.media.array)
      : [];
    const commentMediaDatas = hasPlaceholders(commentXmlData)
      ? getReferencedMedia(commentXmlData, file.media.array)
      : [];
    const footnoteMediaDatas = hasPlaceholders(footnoteXmlData)
      ? getReferencedMedia(footnoteXmlData, file.media.array)
      : [];

    return {
      AppProperties: {
        data: this.formatter.formatToXml(file.appProperties, mkCtx(file.document), {
          encoding: "UTF-8",
          standalone: "yes",
        }),
        path: "docProps/app.xml",
      },
      Comments: {
        data: (() => {
          const xmlData =
            commentMediaDatas.length > 0
              ? replaceImagePlaceholders(
                  commentXmlData,
                  commentMediaDatas,
                  commentRelationshipCount,
                  "rId",
                )
              : commentXmlData;
          return replaceNumberingPlaceholders(xmlData, file.numbering.concreteNumbering);
        })(),
        path: "word/comments.xml",
      },
      CommentsRelationships: {
        data: (() => {
          commentMediaDatas.forEach((mediaData, i) => {
            file.comments.relationships.addRelationship(
              commentRelationshipCount + i,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              `media/${mediaData.fileName}`,
            );
          });
          return this.formatter.formatToXml(
            file.comments.relationships,
            mkCtx({
              relationships: file.comments.relationships,
              view: file.comments,
            }),
            { encoding: "UTF-8" },
          );
        })(),
        path: "word/_rels/comments.xml.rels",
      },
      ContentTypes: {
        data: (() => {
          // Register chart and diagram content types BEFORE serialization
          file.charts.array.forEach((_, i) => {
            file.contentTypes.addChart(i + 1);
          });
          file.smartArts.array.forEach((_, i) => {
            file.contentTypes.addDiagramData(i + 1);
            file.contentTypes.addDiagramLayout(i + 1);
            file.contentTypes.addDiagramStyle(i + 1);
            file.contentTypes.addDiagramColors(i + 1);
            file.contentTypes.addDiagramDrawing(i + 1);
          });
          return this.formatter.formatToXml(file.contentTypes, mkCtx(file.document), {
            encoding: "UTF-8",
          });
        })(),
        path: "[Content_Types].xml",
      },
      CustomProperties: {
        data: this.formatter.formatToXml(file.customProperties, mkCtx(file.document), {
          encoding: "UTF-8",
          standalone: "yes",
        }),
        path: "docProps/custom.xml",
      },
      Document: {
        data: (() => {
          let xmlData =
            documentMediaDatas.length > 0
              ? replaceImagePlaceholders(
                  documentXmlData,
                  documentMediaDatas,
                  documentRelationshipCount,
                  "rId",
                )
              : documentXmlData;
          if (hasPlaceholders(xmlData)) {
            xmlData = replaceChartPlaceholders(
              xmlData,
              file.charts.array.map((c) => c.key),
              documentRelationshipCount + documentMediaDatas.length,
              "rId",
            );
            const smartArtDataOffset =
              documentRelationshipCount + documentMediaDatas.length + file.charts.array.length;
            xmlData = replaceSmartArtPlaceholders(
              xmlData,
              file.smartArts.array.map((s) => s.key),
              smartArtDataOffset,
              "rId",
            );
          }
          const referencedXmlData = replaceNumberingPlaceholders(
            xmlData,
            file.numbering.concreteNumbering,
          );
          return referencedXmlData;
        })(),
        path: "word/document.xml",
      },
      Endnotes: {
        data: this.formatter.formatToXml(file.endnotes.view, mkCtx(file.endnotes), {
          encoding: "UTF-8",
        }),
        path: "word/endnotes.xml",
      },
      EndnotesRelationships: {
        data: this.formatter.formatToXml(file.endnotes.relationships, mkCtx(file.endnotes), {
          encoding: "UTF-8",
        }),
        path: "word/_rels/endnotes.xml.rels",
      },
      FileRelationships: {
        data: this.formatter.formatToXml(file.fileRelationships, mkCtx(file.document), {
          encoding: "UTF-8",
        }),
        path: "_rels/.rels",
      },
      FontTable: {
        data: this.formatter.formatToXml(file.fontTable.view, mkCtx(file.document), {
          encoding: "UTF-8",
          standalone: "yes",
        }),
        path: "word/fontTable.xml",
      },
      FontTableRelationships: {
        data: (() =>
          this.formatter.formatToXml(file.fontTable.relationships, mkCtx(file.document), {
            encoding: "UTF-8",
          }))(),
        path: "word/_rels/fontTable.xml.rels",
      },
      FootNotes: {
        data: (() => {
          const xmlData =
            footnoteMediaDatas.length > 0
              ? replaceImagePlaceholders(
                  footnoteXmlData,
                  footnoteMediaDatas,
                  footnoteRelationshipCount,
                  "rId",
                )
              : footnoteXmlData;
          return replaceNumberingPlaceholders(xmlData, file.numbering.concreteNumbering);
        })(),
        path: "word/footnotes.xml",
      },
      FootNotesRelationships: {
        data: (() => {
          footnoteMediaDatas.forEach((mediaData, i) => {
            file.footNotes.relationships.addRelationship(
              footnoteRelationshipCount + i,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              `media/${mediaData.fileName}`,
            );
          });
          return this.formatter.formatToXml(file.footNotes.relationships, mkCtx(file.footNotes), {
            encoding: "UTF-8",
          });
        })(),
        path: "word/_rels/footnotes.xml.rels",
      },
      FooterRelationships: file.footers.map((footerWrapper, index) => {
        const xmlData = this.formatter.formatToXml(footerWrapper.view, mkCtx(footerWrapper), {
          encoding: "UTF-8",
        });
        // Cache for reuse in Footers section
        footerFormattedViews.set(index, xmlData);
        const mediaDatas = getReferencedMedia(xmlData, file.media.array);

        mediaDatas.forEach((mediaData, i) => {
          footerWrapper.relationships.addRelationship(
            i,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            `media/${mediaData.fileName}`,
          );
        });

        return {
          data: this.formatter.formatToXml(footerWrapper.relationships, mkCtx(footerWrapper), {
            encoding: "UTF-8",
          }),
          path: `word/_rels/footer${index + 1}.xml.rels`,
        };
      }),
      Footers: file.footers.map((_footerWrapper, index) => {
        const tempXmlData = footerFormattedViews.get(index)!;
        const mediaDatas = hasPlaceholders(tempXmlData)
          ? getReferencedMedia(tempXmlData, file.media.array)
          : [];
        const xmlData =
          mediaDatas.length > 0
            ? replaceImagePlaceholders(tempXmlData, mediaDatas, 0, "rId")
            : tempXmlData;

        return {
          data: replaceNumberingPlaceholders(xmlData, file.numbering.concreteNumbering),
          path: `word/footer${index + 1}.xml`,
        };
      }),
      HeaderRelationships: file.headers.map((headerWrapper, index) => {
        const xmlData = this.formatter.formatToXml(headerWrapper.view, mkCtx(headerWrapper), {
          encoding: "UTF-8",
        });
        // Cache for reuse in Headers section
        headerFormattedViews.set(index, xmlData);
        const mediaDatas = getReferencedMedia(xmlData, file.media.array);

        mediaDatas.forEach((mediaData, i) => {
          headerWrapper.relationships.addRelationship(
            i,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            `media/${mediaData.fileName}`,
          );
        });

        return {
          data: this.formatter.formatToXml(headerWrapper.relationships, mkCtx(headerWrapper), {
            encoding: "UTF-8",
          }),
          path: `word/_rels/header${index + 1}.xml.rels`,
        };
      }),
      Headers: file.headers.map((_headerWrapper, index) => {
        const tempXmlData = headerFormattedViews.get(index)!;
        const mediaDatas = hasPlaceholders(tempXmlData)
          ? getReferencedMedia(tempXmlData, file.media.array)
          : [];
        const xmlData =
          mediaDatas.length > 0
            ? replaceImagePlaceholders(tempXmlData, mediaDatas, 0, "rId")
            : tempXmlData;

        return {
          data: replaceNumberingPlaceholders(xmlData, file.numbering.concreteNumbering),
          path: `word/header${index + 1}.xml`,
        };
      }),
      Numbering: {
        data: this.formatter.formatToXml(file.numbering, mkCtx(file.document), {
          encoding: "UTF-8",
          standalone: "yes",
        }),
        path: "word/numbering.xml",
      },
      Properties: {
        data: this.formatter.formatToXml(file.coreProperties, mkCtx(file.document), {
          encoding: "UTF-8",
          standalone: "yes",
        }),
        path: "docProps/core.xml",
      },
      Relationships: {
        data: (() => {
          documentMediaDatas.forEach((mediaData, i) => {
            file.document.relationships.addRelationship(
              documentRelationshipCount + i,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              `media/${mediaData.fileName}`,
            );
          });

          // Chart relationships
          const chartOffset = documentRelationshipCount + documentMediaDatas.length;
          file.charts.array.forEach((_chartData, i) => {
            file.document.relationships.addRelationship(
              chartOffset + i,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
              `charts/chart${i + 1}.xml`,
            );
          });

          // SmartArt relationships (data + layout/style/color internal)
          addSmartArtRelationships(
            file.smartArts.array.map((s) => s.key),
            (id, type, target) => {
              file.document.relationships.addRelationship(id, type, target);
            },
            documentRelationshipCount + documentMediaDatas.length + file.charts.array.length,
            0,
            {
              pathPrefix: "",
              styleRelType: "http://schemas.microsoft.com/office/2007/relationships/diagramStyle",
            },
          );

          file.document.relationships.addRelationship(
            file.document.relationships.relationshipCount + 1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
            "fontTable.xml",
          );

          return this.formatter.formatToXml(file.document.relationships, mkCtx(file.document), {
            encoding: "UTF-8",
          });
        })(),
        path: "word/_rels/document.xml.rels",
      },
      Settings: {
        data: this.formatter.formatToXml(file.settings, mkCtx(file.document), {
          encoding: "UTF-8",
          standalone: "yes",
        }),
        path: "word/settings.xml",
      },
      Styles: {
        data: (() => {
          const xmlStyles = this.formatter.formatToXml(file.styles, mkCtx(file.document), {
            encoding: "UTF-8",
            standalone: "yes",
          });
          const referencedXmlStyles = replaceNumberingPlaceholders(
            xmlStyles,
            file.numbering.concreteNumbering,
          );
          return referencedXmlStyles;
        })(),
        path: "word/styles.xml",
      },
      ...(file.bibliography
        ? {
            Bibliography: {
              data: this.formatter.formatToXml(
                file.bibliography,
                mkCtx({
                  relationships: file.bibliography.relationships,
                  view: file.bibliography,
                }),
                { encoding: "UTF-8", standalone: "yes" },
              ),
              path: "word/bibliography.xml",
            },
          }
        : {}),
      ...(file.charts.array.length > 0
        ? {
            Charts: file.charts.array.flatMap((chartData, i) => [
              {
                data: this.formatter.formatToXml(chartData.chartSpace, mkCtx(file.document), {
                  encoding: "UTF-8",
                  standalone: "yes",
                }),
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
      ...(file.smartArts.array.length > 0
        ? {
            DiagramData: file.smartArts.array.map((smartArtData, i) => ({
              data: this.formatter.formatToXml(smartArtData.dataModel, mkCtx(file.document), {
                encoding: "UTF-8",
                standalone: "yes",
              }),
              path: `word/diagrams/data${i + 1}.xml`,
            })),
            DiagramLayout: file.smartArts.array.map((smartArtData, i) => ({
              data: getLayoutXml(smartArtData.layout),
              path: `word/diagrams/layout${i + 1}.xml`,
            })),
            DiagramStyle: file.smartArts.array.map((smartArtData, i) => ({
              data: getStyleXml(smartArtData.style),
              path: `word/diagrams/quickStyle${i + 1}.xml`,
            })),
            DiagramColors: file.smartArts.array.map((smartArtData, i) => ({
              data: getColorXml(smartArtData.color),
              path: `word/diagrams/colors${i + 1}.xml`,
            })),
            DiagramDrawing: file.smartArts.array.map((_, i) => ({
              data: DEFAULT_DRAWING_XML,
              path: `word/diagrams/drawing${i + 1}.xml`,
            })),
          }
        : {}),
      ...(file.altChunks.array.length > 0
        ? {
            AltChunks: file.altChunks.array.map((altChunkData) => ({
              data: altChunkData.data,
              path: `word/${altChunkData.path}`,
            })),
          }
        : {}),
      ...(file.subDocs.array.length > 0
        ? {
            SubDocs: file.subDocs.array.map((subDocData) => ({
              data: subDocData.data,
              path: `word/${subDocData.path}`,
            })),
          }
        : {}),
    };
  }
}
