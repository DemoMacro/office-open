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
  compileMapping,
  getReferencedMedia,
  hasPlaceholders,
  replaceChartPlaceholders,
  replaceImagePlaceholders,
  replaceSmartArtPlaceholders,
} from "@office-open/core";
import type { XmlifyedFile, Zippable } from "@office-open/core";
import { xml } from "@office-open/xml";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

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

  public compile(file: File, overrides: readonly XmlifyedFile[] = []): Zippable {
    const declaration = true;
    const context: Context = { fileData: file, stack: [] };

    const mapping: XmlifyedFileMapping = {
      AppProperties: {
        data: this.formatter.formatToXml(file.appProperties, context, declaration),
        path: "docProps/app.xml",
      },
      Properties: {
        data: this.formatter.formatToXml(file.coreProperties, context, declaration),
        path: "docProps/core.xml",
      },
      FileRelationships: {
        data: this.formatter.formatToXml(file.fileRelationships, context),
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
        data: this.formatter.formatToXml(masterRels[mi], context),
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
        data: this.formatter.formatToXml(layoutRels[li], context),
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
        data: this.formatter.formatToXml(new DefaultNotesMaster(), context, declaration),
        path: "ppt/notesMasters/notesMaster1.xml",
      };
      mapping["NotesMasterRelationships"] = {
        data: this.formatter.formatToXml(file.notesMasterRelationships, context),
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
    for (let idx = 0; idx < mediaData.length; idx++) {
      file.presentationWrapper.relationships.addRelationship(
        presImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${mediaData[idx].fileName}`,
      );
    }

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
      data: this.formatter.formatToXml(file.presentationWrapper.relationships, context),
      path: "ppt/_rels/presentation.xml.rels",
    };

    // Slides — format BEFORE ContentTypes so Chart.prepForXml() populates Charts
    for (let i = 0; i < file.slides.length; i++) {
      const slideWrapper = file.slideWrappers[i];
      const slideXml = this.formatter.formatToXml(slideWrapper.view, context, declaration);

      const slideMediaData = getReferencedMedia(slideXml, file.media.array);
      const slideImageOffset = slideWrapper.relationships.relationshipCount + 1;
      for (let idx = 0; idx < slideMediaData.length; idx++) {
        slideWrapper.relationships.addRelationship(
          slideImageOffset + idx,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          `../media/${slideMediaData[idx].fileName}`,
        );
      }

      let replacedSlideXml = replaceImagePlaceholders(slideXml, slideMediaData, slideImageOffset);
      currentImageCount += slideMediaData.length;

      // Placeholder replacement — single pre-check avoids 4 regex scans when none exist
      if (hasPlaceholders(replacedSlideXml)) {
        // Chart placeholder replacement
        const slideChartKeys = collectPlaceholderKeys(replacedSlideXml, "chart:");
        if (slideChartKeys.length > 0) {
          const slideChartOffset = slideWrapper.relationships.relationshipCount + 1;
          const slideChartKeySet = new Set(slideChartKeys);
          const slideCharts = file.charts.array.filter((c) => slideChartKeySet.has(c.key));

          replacedSlideXml = replaceChartPlaceholders(
            replacedSlideXml,
            slideCharts.map((c) => c.key),
            slideChartOffset,
          );

          for (let ci = 0; ci < slideCharts.length; ci++) {
            const chartData = slideCharts[ci];
            const globalIndex = file.charts.array.indexOf(chartData);
            slideWrapper.relationships.addRelationship(
              slideChartOffset + ci,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
              `../charts/chart${globalIndex + 1}.xml`,
            );
          }
        }

        // SmartArt placeholder replacement
        const slideSmartArtKeys = collectPlaceholderKeys(replacedSlideXml, "smartart:");
        if (slideSmartArtKeys.length > 0) {
          const slideSmartArtKeySet = new Set(slideSmartArtKeys);
          const slideSmartArts = file.smartArts.array.filter((s) => slideSmartArtKeySet.has(s.key));
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
          const slideHlinkKeySet = new Set(slideHlinkKeys);
          const slideHlinks = file.hyperlinks.array.filter((h) => slideHlinkKeySet.has(h.key));
          const hlinkOffset = slideWrapper.relationships.relationshipCount + 1;

          replacedSlideXml = replaceHyperlinkPlaceholders(
            replacedSlideXml,
            slideHlinks,
            hlinkOffset,
          );

          for (let hi = 0; hi < slideHlinks.length; hi++) {
            slideWrapper.relationships.addRelationship(
              hlinkOffset + hi,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
              slideHlinks[hi].url,
              "External",
            );
          }
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

          for (let mi = 0; mi < slideMediaRefs.length; mi++) {
            slideWrapper.relationships.addRelationship(
              mediaOffset + mi,
              "http://schemas.microsoft.com/office/2007/relationships/media",
              `../media/${slideMediaRefs[mi].fileName}`,
            );
          }

          for (let vi = 0; vi < slideVideoRefs.length; vi++) {
            slideWrapper.relationships.addRelationship(
              videoOffset + vi,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
              `../media/${slideVideoRefs[vi].fileName}`,
            );
          }
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
        data: this.formatter.formatToXml(slideWrapper.relationships, context),
        path: `ppt/slides/_rels/slide${i + 1}.xml.rels`,
      };
    }

    // ContentTypes — AFTER slides so Charts.Array is populated
    for (let i = 0; i < file.charts.array.length; i++) {
      file.contentTypes.addChart(i + 1);
    }
    for (let i = 0; i < file.smartArts.array.length; i++) {
      file.contentTypes.addDiagramData(i + 1);
      file.contentTypes.addDiagramLayout(i + 1);
      file.contentTypes.addDiagramStyle(i + 1);
      file.contentTypes.addDiagramColors(i + 1);
      file.contentTypes.addDiagramDrawing(i + 1);
    }
    mapping["ContentTypes"] = {
      data: this.formatter.formatToXml(file.contentTypes, context),
      path: "[Content_Types].xml",
    };

    // Convert mapping to Zippable (XML files use global DEFLATE default)
    const files = compileMapping(mapping, overrides);

    // Add chart parts
    for (let i = 0; i < file.charts.array.length; i++) {
      const chartData = file.charts.array[i];
      files[`ppt/charts/chart${i + 1}.xml`] = encoder.encode(
        this.formatter.formatToXml(chartData.chartSpace, context, declaration),
      );
      files[`ppt/charts/_rels/chart${i + 1}.xml.rels`] = encoder.encode(
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
      files[`ppt/diagrams/data${i + 1}.xml`] = encoder.encode(
        this.formatter.formatToXml(smartArtData.dataModel, context, declaration),
      );
      files[`ppt/diagrams/layout${i + 1}.xml`] = encoder.encode(getLayoutXml(smartArtData.layout));
      files[`ppt/diagrams/quickStyle${i + 1}.xml`] = encoder.encode(
        getStyleXml(smartArtData.style),
      );
      files[`ppt/diagrams/colors${i + 1}.xml`] = encoder.encode(getColorXml(smartArtData.color));
      files[`ppt/diagrams/drawing${i + 1}.xml`] = encoder.encode(DEFAULT_DRAWING_XML);
    }

    // Add notes slides
    for (let i = 0; i < file.notesSlides.length; i++) {
      const notesSlide = file.notesSlides[i];
      files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = encoder.encode(
        this.formatter.formatToXml(notesSlide, context, declaration),
      );
      files[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = encoder.encode(
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
      files["ppt/commentAuthors.xml"] = encoder.encode(
        this.formatter.formatToXml(file.commentAuthorList, context, declaration),
      );
    }

    // Add slide comments
    const commentLists = file.slideCommentLists;
    for (let i = 0; i < commentLists.length; i++) {
      if (commentLists[i]) {
        files[`ppt/comments/comment${i + 1}.xml`] = encoder.encode(
          this.formatter.formatToXml(commentLists[i]!, context, declaration),
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
