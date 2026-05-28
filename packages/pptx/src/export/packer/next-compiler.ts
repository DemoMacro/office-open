import { Formatter } from "@export/formatter";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { DefaultNotesMaster } from "@file/notes-master/notes-master";
import {
  DEFAULT_DRAWING_XML,
  getColorXml,
  getLayoutXml,
  getStyleXml,
} from "@file/smartart/built-in-definitions";
import type { Context } from "@file/xml-components";
import {
  ZIP_STORED_LEVEL,
  addSmartArtRelationships,
  collectPlaceholderKeys,
  getReferencedMedia,
  hasPlaceholders,
  replaceChartPlaceholders,
  replaceImagePlaceholders,
  replaceSmartArtPlaceholders,
} from "@office-open/core";
import type { XmlifyedFile, Zippable } from "@office-open/core";
import { xml } from "@office-open/xml";

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
        data: xml(this.formatter.format(file.appProperties, context), { declaration }),
        path: "docProps/app.xml",
      },
      Properties: {
        data: xml(this.formatter.format(file.coreProperties, context), {
          declaration,
          indent,
        }),
        path: "docProps/core.xml",
      },
      FileRelationships: {
        data: xml(this.formatter.format(file.fileRelationships, context), {
          declaration: false,
        }),
        path: "_rels/.rels",
      },
    };

    // Static singletons (ImportedXmlComponent.toXml uses sourceXml directly)
    // Themes (one per master)
    const themes = file.themes;
    for (let ti = 0; ti < themes.length; ti++) {
      mapping[`Theme${ti}`] = {
        data: this.formatter.formatToXml(themes[ti], context, declaration),
        path: `ppt/theme/theme${ti + 1}.xml`,
      };
    }

    mapping["TableStyles"] = {
      data: this.formatter.formatToXml(file.tableStyles, context, declaration),
      path: "ppt/tableStyles.xml",
    };

    mapping["PresProps"] = {
      data: this.formatter.formatToXml(file.presProps, context, declaration),
      path: "ppt/presProps.xml",
    };

    mapping["ViewProps"] = {
      data: this.formatter.formatToXml(file.viewProps, context, declaration),
      path: "ppt/viewProps.xml",
    };

    // Slide Masters
    const masters = file.slideMasters;
    const masterRels = file.slideMasterRelsArray;
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
    const layouts = file.allLayouts;
    const layoutRels = file.allLayoutRelsArray;
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
      file.contentTypes.addSlideLayout(i + 1);
    }
    // Register master + theme content types
    for (let mi = 0; mi < masters.length; mi++) {
      file.contentTypes.addSlideMaster(mi + 1);
      file.contentTypes.addTheme(mi + 1);
    }

    // Notes Master — only when notes slides exist
    if (file.notesSlides.length > 0) {
      file.presentationWrapper.relationships.addRelationship(
        file.presentationWrapper.relationships.relationshipCount + 1,
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
        data: xml(this.formatter.format(file.notesMasterRelationships, context), {
          declaration: false,
        }),
        path: "ppt/notesMasters/_rels/notesMaster1.xml.rels",
      };
    }

    // Comment Authors — only when comments exist
    if (file.commentAuthorList) {
      file.presentationWrapper.relationships.addRelationship(
        file.presentationWrapper.relationships.relationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
        "commentAuthors.xml",
      );
    }

    // Presentation + its relationships
    const presentationXml = this.formatter.formatToXml(
      file.presentationWrapper.view,
      context,
      declaration,
    );
    let currentImageCount = 0;

    const mediaData = getReferencedMedia(presentationXml, file.media.array);
    const presImageOffset = file.presentationWrapper.relationships.relationshipCount + 1;
    mediaData.forEach((image, idx) => {
      file.presentationWrapper.relationships.addRelationship(
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
      data: xml(this.formatter.format(file.presentationWrapper.relationships, context), {
        declaration: false,
      }),
      path: "ppt/_rels/presentation.xml.rels",
    };

    // Slides — format BEFORE ContentTypes so ChartFrame.prepForXml() populates Charts
    for (let i = 0; i < file.slides.length; i++) {
      const slideWrapper = file.slideWrappers[i];
      const slideXml = this.formatter.formatToXml(slideWrapper.view, context, declaration);

      const slideMediaData = getReferencedMedia(slideXml, file.media.array);
      const slideImageOffset = slideWrapper.relationships.relationshipCount + 1;
      slideMediaData.forEach((image, idx) => {
        slideWrapper.relationships.addRelationship(
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
          const slideChartOffset = slideWrapper.relationships.relationshipCount + 1;
          const slideCharts = file.charts.array.filter((c) => slideChartKeys.includes(c.key));

          replacedSlideXml = replaceChartPlaceholders(
            replacedSlideXml,
            slideCharts.map((c) => c.key),
            slideChartOffset,
          );

          slideCharts.forEach((chartData, ci) => {
            const globalIndex = file.charts.array.indexOf(chartData);
            slideWrapper.relationships.addRelationship(
              slideChartOffset + ci,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
              `../charts/chart${globalIndex + 1}.xml`,
            );
          });
        }

        // SmartArt placeholder replacement
        const slideSmartArtKeys = collectPlaceholderKeys(replacedSlideXml, "smartart:");
        if (slideSmartArtKeys.length > 0) {
          const slideSmartArts = file.smartArts.array.filter((s) =>
            slideSmartArtKeys.includes(s.key),
          );
          const saOffset = slideWrapper.relationships.relationshipCount + 1;

          replacedSlideXml = replaceSmartArtPlaceholders(
            replacedSlideXml,
            slideSmartArts.map((s) => s.key),
            saOffset,
          );

          const saGlobalStart = file.smartArts.array.indexOf(slideSmartArts[0]);
          addSmartArtRelationships(
            slideSmartArts.map((s) => s.key),
            (id, type, target) => {
              slideWrapper.relationships.addRelationship(id, type, target);
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
          const slideHlinks = file.hyperlinks.array.filter((h) => slideHlinkKeys.includes(h.key));
          const hlinkOffset = slideWrapper.relationships.relationshipCount + 1;

          replacedSlideXml = replaceHyperlinkPlaceholders(
            replacedSlideXml,
            slideHlinks,
            hlinkOffset,
          );

          slideHlinks.forEach((hlink, hi) => {
            slideWrapper.relationships.addRelationship(
              hlinkOffset + hi,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
              hlink.url,
              "External",
            );
          });
        }

        // Media (video/audio) placeholder replacement
        const slideMediaRefs = getMediaRefs(replacedSlideXml, file.media.array);
        const slideVideoRefs = getVideoRefs(replacedSlideXml, file.media.array);

        if (slideMediaRefs.length > 0 || slideVideoRefs.length > 0) {
          const mediaOffset = slideWrapper.relationships.relationshipCount + 1;
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
            slideWrapper.relationships.addRelationship(
              mediaOffset + mi,
              "http://schemas.microsoft.com/office/2007/relationships/media",
              `../media/${media.fileName}`,
            );
          });

          slideVideoRefs.forEach((video, vi) => {
            slideWrapper.relationships.addRelationship(
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
      const slideComments = file.slideCommentLists[i];
      if (slideComments) {
        slideWrapper.relationships.addRelationship(
          slideWrapper.relationships.relationshipCount + 1,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
          `../comments/comment${i + 1}.xml`,
        );
      }

      // Add notesSlide relationship for this slide (if it has notes)
      const notesSlideIndex = file.notesSlideIndexMap.get(i);
      if (notesSlideIndex !== undefined) {
        slideWrapper.relationships.addRelationship(
          slideWrapper.relationships.relationshipCount + 1,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
          `../notesSlides/notesSlide${notesSlideIndex + 1}.xml`,
        );
      }

      mapping[`SlideRelationships${i}`] = {
        data: xml(this.formatter.format(slideWrapper.relationships, context), {
          declaration: false,
        }),
        path: `ppt/slides/_rels/slide${i + 1}.xml.rels`,
      };
    }

    // ContentTypes — AFTER slides so Charts.Array is populated
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
    mapping["ContentTypes"] = {
      data: xml(this.formatter.format(file.contentTypes, context), {
        declaration: false,
      }),
      path: "[Content_Types].xml",
    };

    // Convert mapping to Zippable (XML files use global DEFLATE default)
    const files: Zippable = {};
    for (const key of Object.keys(mapping)) {
      const entry = mapping[key];
      files[entry.path] = textToUint8Array(entry.data);
    }

    // Add overrides
    for (const override of overrides) {
      files[override.path] =
        override.data instanceof Uint8Array ? override.data : textToUint8Array(override.data);
    }

    // Add chart parts
    for (let i = 0; i < file.charts.array.length; i++) {
      const chartData = file.charts.array[i];
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

    // Add SmartArt diagram parts
    for (let i = 0; i < file.smartArts.array.length; i++) {
      const smartArtData = file.smartArts.array[i];
      files[`ppt/diagrams/data${i + 1}.xml`] = textToUint8Array(
        xml(this.formatter.format(smartArtData.dataModel, context), {
          declaration,
          indent,
        }),
      );
      files[`ppt/diagrams/layout${i + 1}.xml`] = textToUint8Array(
        getLayoutXml(smartArtData.layout),
      );
      files[`ppt/diagrams/quickStyle${i + 1}.xml`] = textToUint8Array(
        getStyleXml(smartArtData.style),
      );
      files[`ppt/diagrams/colors${i + 1}.xml`] = textToUint8Array(getColorXml(smartArtData.color));
      files[`ppt/diagrams/drawing${i + 1}.xml`] = textToUint8Array(DEFAULT_DRAWING_XML);
    }

    // Add notes slides
    for (let i = 0; i < file.notesSlides.length; i++) {
      const notesSlide = file.notesSlides[i];
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

    // Add comment authors
    if (file.commentAuthorList) {
      files["ppt/commentAuthors.xml"] = textToUint8Array(
        xml(this.formatter.format(file.commentAuthorList, context), {
          declaration,
          indent,
        }),
      );
    }

    // Add slide comments
    const commentLists = file.slideCommentLists;
    for (let i = 0; i < commentLists.length; i++) {
      if (commentLists[i]) {
        files[`ppt/comments/comment${i + 1}.xml`] = textToUint8Array(
          xml(this.formatter.format(commentLists[i]!, context), {
            declaration,
            indent,
          }),
        );
      }
    }

    // Add media files (STORE — already-compressed formats)
    for (const image of file.media.array) {
      files[`ppt/media/${image.fileName}`] = [image.data, { level: ZIP_STORED_LEVEL }];
      if (image.type === "svg" && "fallback" in image) {
        const fallback = (
          image as IMediaData & {
            readonly fallback: { readonly fileName: string; readonly data: Uint8Array };
          }
        ).fallback;
        files[`ppt/media/${fallback.fileName}`] = [fallback.data, { level: ZIP_STORED_LEVEL }];
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
