import { Formatter } from "@export/formatter";
import type { File } from "@file/file";
import { DefaultNotesMaster } from "@file/notes-master/notes-master";
import {
  DEFAULT_DRAWING_XML,
  getColorXml,
  getLayoutXml,
  getStyleXml,
} from "@file/smartart/built-in-definitions";
import type { Context } from "@file/xml-components";
import {
  addSmartArtRelationships,
  collectPlaceholderKeys,
  getReferencedMedia,
  hasPlaceholders,
  replaceChartPlaceholders,
  replaceImagePlaceholders,
  replaceSmartArtPlaceholders,
} from "@office-open/core";
import type { XmlifyedFile } from "@office-open/core";
import { xml } from "@office-open/xml";
import type { Zippable } from "fflate";

interface XmlifyedFileMapping {
  [key: string]: { data: string; path: string };
}

import { replaceHyperlinkPlaceholders } from "./hyperlink-placeholders";
import {
  getMediaRefs,
  getVideoRefs,
  replaceMediaPlaceholders,
  replaceVideoPlaceholders,
} from "./media-placeholders";

export class Compiler {
  private readonly formatter = new Formatter();

  public compile(
    file: File,
    prettifyXml?: string,
    overrides: readonly XmlifyedFile[] = [],
  ): Zippable {
    const declaration = true;
    const indent = prettifyXml;
    const context: Context = { fileData: file, stack: [] };

    const mapping: XmlifyedFileMapping = {
      AppProperties: {
        data: xml(this.formatter.format(file.AppProperties, context), { declaration }),
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

    // Static singletons (ImportedXmlComponent.toXml uses sourceXml directly)
    // Themes (one per master)
    const themes = file.Themes;
    for (let ti = 0; ti < themes.length; ti++) {
      mapping[`Theme${ti}`] = {
        data: this.formatter.formatToXml(themes[ti], context, declaration),
        path: `ppt/theme/theme${ti + 1}.xml`,
      };
    }

    mapping["TableStyles"] = {
      data: this.formatter.formatToXml(file.TableStyles, context, declaration),
      path: "ppt/tableStyles.xml",
    };

    mapping["PresProps"] = {
      data: this.formatter.formatToXml(file.PresProps, context, declaration),
      path: "ppt/presProps.xml",
    };

    mapping["ViewProps"] = {
      data: this.formatter.formatToXml(file.ViewProps, context, declaration),
      path: "ppt/viewProps.xml",
    };

    // Slide Masters
    const masters = file.SlideMasters;
    const masterRels = file.SlideMasterRelsArray;
    for (let mi = 0; mi < masters.length; mi++) {
      mapping[`SlideMaster${mi}`] = {
        data: this.formatter.formatToXml(masters[mi], context, declaration),
        path: `ppt/slideMasters/slideMaster${mi + 1}.xml`,
      };
      mapping[`SlideMasterRels${mi}`] = {
        data: xml(this.formatter.format(masterRels[mi], context), {
          declaration: false,
        }),
        path: `ppt/slideMasters/_rels/slideMaster${mi + 1}.xml.rels`,
      };
    }

    // Slide Layouts
    const layouts = file.AllLayouts;
    const layoutRels = file.AllLayoutRelsArray;
    for (let li = 0; li < layouts.length; li++) {
      mapping[`SlideLayout${li}`] = {
        data: this.formatter.formatToXml(layouts[li].layout, context, declaration),
        path: `ppt/slideLayouts/slideLayout${li + 1}.xml`,
      };
      mapping[`SlideLayoutRels${li}`] = {
        data: xml(this.formatter.format(layoutRels[li], context), {
          declaration: false,
        }),
        path: `ppt/slideLayouts/_rels/slideLayout${li + 1}.xml.rels`,
      };
    }

    // Register layout content types
    for (let i = 0; i < layouts.length; i++) {
      file.ContentTypes.addSlideLayout(i + 1);
    }
    // Register master + theme content types
    for (let mi = 0; mi < masters.length; mi++) {
      file.ContentTypes.addSlideMaster(mi + 1);
      file.ContentTypes.addTheme(mi + 1);
    }

    // Notes Master — only when notes slides exist
    if (file.NotesSlides.length > 0) {
      file.PresentationWrapper.Relationships.addRelationship(
        file.PresentationWrapper.Relationships.RelationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
        "notesMasters/notesMaster1.xml",
      );
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
    }

    // Comment Authors — only when comments exist
    if (file.CommentAuthorList) {
      file.PresentationWrapper.Relationships.addRelationship(
        file.PresentationWrapper.Relationships.RelationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
        "commentAuthors.xml",
      );
    }

    // Presentation + its relationships
    const presentationXml = this.formatter.formatToXml(
      file.PresentationWrapper.View,
      context,
      declaration,
    );
    let currentImageCount = 0;

    const mediaData = getReferencedMedia(presentationXml, file.Media.Array);
    const presImageOffset = file.PresentationWrapper.Relationships.RelationshipCount + 1;
    mediaData.forEach((image, idx) => {
      file.PresentationWrapper.Relationships.addRelationship(
        presImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${image.fileName}`,
      );
    });

    const replacedPresentationXml = replaceImagePlaceholders(
      presentationXml,
      mediaData,
      presImageOffset,
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
      const slideXml = this.formatter.formatToXml(slideWrapper.View, context, declaration);

      const slideMediaData = getReferencedMedia(slideXml, file.Media.Array);
      const slideImageOffset = slideWrapper.Relationships.RelationshipCount + 1;
      slideMediaData.forEach((image, idx) => {
        slideWrapper.Relationships.addRelationship(
          slideImageOffset + idx,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          `../media/${image.fileName}`,
        );
      });

      let replacedSlideXml = replaceImagePlaceholders(slideXml, slideMediaData, slideImageOffset);
      currentImageCount += slideMediaData.length;

      // Placeholder replacement — single pre-check avoids 4 regex scans when none exist
      if (hasPlaceholders(replacedSlideXml)) {
        // Chart placeholder replacement
        const slideChartKeys = collectPlaceholderKeys(replacedSlideXml, "chart:");
        if (slideChartKeys.length > 0) {
          const slideChartOffset = slideWrapper.Relationships.RelationshipCount + 1;
          const slideCharts = file.Charts.Array.filter((c) => slideChartKeys.includes(c.key));

          replacedSlideXml = replaceChartPlaceholders(
            replacedSlideXml,
            slideCharts.map((c) => c.key),
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

        // SmartArt placeholder replacement
        const slideSmartArtKeys = collectPlaceholderKeys(replacedSlideXml, "smartart:");
        if (slideSmartArtKeys.length > 0) {
          const slideSmartArts = file.SmartArts.Array.filter((s) =>
            slideSmartArtKeys.includes(s.key),
          );
          const saOffset = slideWrapper.Relationships.RelationshipCount + 1;

          replacedSlideXml = replaceSmartArtPlaceholders(
            replacedSlideXml,
            slideSmartArts.map((s) => s.key),
            saOffset,
          );

          const saGlobalStart = file.SmartArts.Array.indexOf(slideSmartArts[0]);
          addSmartArtRelationships(
            slideSmartArts.map((s) => s.key),
            (id, type, target) => {
              slideWrapper.Relationships.addRelationship(id, type, target);
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

        // Hyperlink placeholder replacement
        const slideHlinkKeys = collectPlaceholderKeys(replacedSlideXml, "hlink:");
        if (slideHlinkKeys.length > 0) {
          const slideHlinks = file.Hyperlinks.Array.filter((h) => slideHlinkKeys.includes(h.key));
          const hlinkOffset = slideWrapper.Relationships.RelationshipCount + 1;

          replacedSlideXml = replaceHyperlinkPlaceholders(
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
        const slideMediaRefs = getMediaRefs(replacedSlideXml, file.Media.Array);
        const slideVideoRefs = getVideoRefs(replacedSlideXml, file.Media.Array);

        if (slideMediaRefs.length > 0 || slideVideoRefs.length > 0) {
          const mediaOffset = slideWrapper.Relationships.RelationshipCount + 1;
          const videoOffset = mediaOffset + slideMediaRefs.length;

          replacedSlideXml = replaceMediaPlaceholders(
            replacedSlideXml,
            slideMediaRefs,
            mediaOffset,
          );

          replacedSlideXml = replaceVideoPlaceholders(
            replacedSlideXml,
            slideVideoRefs,
            videoOffset,
          );

          slideMediaRefs.forEach((media, mi) => {
            slideWrapper.Relationships.addRelationship(
              mediaOffset + mi,
              "http://schemas.microsoft.com/office/2007/relationships/media",
              `../media/${media.fileName}`,
            );
          });

          slideVideoRefs.forEach((video, vi) => {
            slideWrapper.Relationships.addRelationship(
              videoOffset + vi,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
              `../media/${video.fileName}`,
            );
          });
        }
      }

      mapping[`Slide${i}`] = {
        data: replacedSlideXml,
        path: `ppt/slides/slide${i + 1}.xml`,
      };

      // Add comment relationship for this slide
      const slideComments = file.SlideCommentLists[i];
      if (slideComments) {
        slideWrapper.Relationships.addRelationship(
          slideWrapper.Relationships.RelationshipCount + 1,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
          `../comments/comment${i + 1}.xml`,
        );
      }

      // Add notesSlide relationship for this slide (if it has notes)
      const notesSlideIndex = file.NotesSlideIndexMap.get(i);
      if (notesSlideIndex !== undefined) {
        slideWrapper.Relationships.addRelationship(
          slideWrapper.Relationships.RelationshipCount + 1,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
          `../notesSlides/notesSlide${notesSlideIndex + 1}.xml`,
        );
      }

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
    file.SmartArts.Array.forEach((_, i) => {
      file.ContentTypes.addDiagramData(i + 1);
      file.ContentTypes.addDiagramLayout(i + 1);
      file.ContentTypes.addDiagramStyle(i + 1);
      file.ContentTypes.addDiagramColors(i + 1);
      file.ContentTypes.addDiagramDrawing(i + 1);
    });
    mapping["ContentTypes"] = {
      data: xml(this.formatter.format(file.ContentTypes, context), {
        declaration: false,
      }),
      path: "[Content_Types].xml",
    };

    // Convert mapping to Zippable (XML files: STORE, no compression benefit)
    const files: Zippable = {};
    for (const key of Object.keys(mapping)) {
      const entry = mapping[key];
      files[entry.path] = [textToUint8Array(entry.data), { level: 0 }];
    }

    // Add overrides
    for (const override of overrides) {
      files[override.path] =
        override.data instanceof Uint8Array ? override.data : textToUint8Array(override.data);
    }

    // Add chart parts (STORE — XML, no compression benefit)
    for (let i = 0; i < file.Charts.Array.length; i++) {
      const chartData = file.Charts.Array[i];
      files[`ppt/charts/chart${i + 1}.xml`] = [
        textToUint8Array(
          xml(this.formatter.format(chartData.chartSpace, context), {
            declaration,
            indent,
          }),
        ),
        { level: 0 },
      ];
      files[`ppt/charts/_rels/chart${i + 1}.xml.rels`] = [
        textToUint8Array(
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
        ),
        { level: 0 },
      ];
    }

    // Add SmartArt diagram parts (STORE)
    for (let i = 0; i < file.SmartArts.Array.length; i++) {
      const smartArtData = file.SmartArts.Array[i];
      files[`ppt/diagrams/data${i + 1}.xml`] = [
        textToUint8Array(
          xml(this.formatter.format(smartArtData.dataModel, context), {
            declaration,
            indent,
          }),
        ),
        { level: 0 },
      ];
      files[`ppt/diagrams/layout${i + 1}.xml`] = [
        textToUint8Array(getLayoutXml(smartArtData.layout)),
        { level: 0 },
      ];
      files[`ppt/diagrams/quickStyle${i + 1}.xml`] = [
        textToUint8Array(getStyleXml(smartArtData.style)),
        { level: 0 },
      ];
      files[`ppt/diagrams/colors${i + 1}.xml`] = [
        textToUint8Array(getColorXml(smartArtData.color)),
        { level: 0 },
      ];
      files[`ppt/diagrams/drawing${i + 1}.xml`] = [
        textToUint8Array(DEFAULT_DRAWING_XML),
        { level: 0 },
      ];
    }

    // Add notes slides (STORE)
    for (let i = 0; i < file.NotesSlides.length; i++) {
      const notesSlide = file.NotesSlides[i];
      files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = [
        textToUint8Array(
          xml(this.formatter.format(notesSlide, context), {
            declaration,
            indent,
          }),
        ),
        { level: 0 },
      ];
      files[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = [
        textToUint8Array(
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
        ),
        { level: 0 },
      ];
    }

    // Add comment authors (STORE)
    if (file.CommentAuthorList) {
      files["ppt/commentAuthors.xml"] = [
        textToUint8Array(
          xml(this.formatter.format(file.CommentAuthorList, context), {
            declaration,
            indent,
          }),
        ),
        { level: 0 },
      ];
    }

    // Add slide comments (STORE)
    const commentLists = file.SlideCommentLists;
    for (let i = 0; i < commentLists.length; i++) {
      if (commentLists[i]) {
        files[`ppt/comments/comment${i + 1}.xml`] = [
          textToUint8Array(
            xml(this.formatter.format(commentLists[i]!, context), {
              declaration,
              indent,
            }),
          ),
          { level: 0 },
        ];
      }
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
