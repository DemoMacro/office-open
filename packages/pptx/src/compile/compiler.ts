/**
 * Descriptor-based PPTX compiler.
 *
 * Produces a valid PPTX ZIP archive using the new descriptor pipeline
 * for slide serialization, while reusing File for structural setup
 * (themes, masters, layouts, relationships, content types).
 *
 * @module
 */

import { replaceHyperlinkPlaceholders } from "@export/packer/hyperlink-placeholders";
import {
  getMediaRefs,
  getVideoRefs,
  replaceMediaPlaceholders,
  replaceVideoPlaceholders,
} from "@export/packer/media-placeholders";
import { SP_TREE_HEADER } from "@file/constants";
import type { File, SlideOptions } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { getColorXml, getLayoutXml, getStyleXml, DEFAULT_DRAWING_XML } from "@file/smartart";
import { DefaultTheme } from "@file/theme";
import { buildTransition } from "@file/transition";
import { Relationships } from "@office-open/core";
import {
  APP_PROPS_XML,
  buildCorePropertiesXmlString,
  collectPlaceholderKeys,
  compileMapping,
  getReferencedMedia,
  hasPlaceholders,
  replaceChartPlaceholders,
  replaceImagePlaceholders,
  replaceSmartArtPlaceholders,
  addSmartArtRelationships,
} from "@office-open/core";
import type { XmlifyedFile, ZipOptions, Zippable } from "@office-open/core";

import { PptxWriteContext } from "./context";
import { timingDesc } from "./descriptors/animation";
import { backgroundDesc } from "./descriptors/background";
import { stringifyChild } from "./descriptors/bridge";
import { commentAuthorsDesc, slideCommentsDesc } from "./descriptors/comments";
import { contentTypesDesc } from "./descriptors/content-types";
import { handoutMasterDesc } from "./descriptors/handout-master";
import { notesMasterDesc } from "./descriptors/notes-master";
import { notesSlideDesc } from "./descriptors/notes-slide";
import { presentationDesc } from "./descriptors/presentation";
import { presPropsDesc } from "./descriptors/presentation-properties";
import { slideLayoutDesc } from "./descriptors/slide-layout";
import { slideMasterDesc } from "./descriptors/slide-master";
import { slideSyncDesc } from "./descriptors/slide-sync";
import { tableStylesDesc } from "./descriptors/table-styles";
import { viewPropsDesc } from "./descriptors/view-properties";

/** Reusable TextEncoder. */
const encoder = new TextEncoder();

/** XML declaration prefix. */
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

interface XmlifyedFileMapping {
  [key: string]: { data: string; path: string };
}

// ── Slide serializer using descriptors ──

function stringifySlide(slideOpts: SlideOptions, ctx: PptxWriteContext): string {
  const parts: string[] = [];

  // Opening tag
  const sldAttrs: string[] = [];
  if (slideOpts.showMasterShapes === false) sldAttrs.push(' showMasterSp="0"');
  if (slideOpts.showMasterPlaceholderAnimations === false) sldAttrs.push(' showMasterPhAnim="0"');
  parts.push(
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${sldAttrs.join("")}>`,
  );

  // p:cSld
  parts.push("<p:cSld>");

  if (slideOpts.background) {
    parts.push(backgroundDesc.stringify(slideOpts.background, ctx) ?? "");
  }

  // p:spTree
  parts.push("<p:spTree>");
  parts.push(SP_TREE_HEADER);

  if (slideOpts.children) {
    for (const child of slideOpts.children) {
      const xml = stringifyChild(child, ctx);
      if (xml) parts.push(xml);
    }
  }

  parts.push("</p:spTree>");

  // custDataLst
  if (slideOpts.customerData && slideOpts.customerData.length > 0) {
    const cdItems = slideOpts.customerData.map((d) => `<p:custData r:id="${d.rId}"/>`).join("");
    parts.push(`<p:custDataLst>${cdItems}</p:custDataLst>`);
  }

  // controls
  if (slideOpts.controls && slideOpts.controls.length > 0) {
    const ctrlItems = slideOpts.controls
      .map((c) => {
        const attrs: string[] = [];
        if (c.shapeId !== undefined) attrs.push(`spid="${c.shapeId}"`);
        if (c.name) attrs.push(`name="${c.name}"`);
        if (c.showAsIcon) attrs.push('showAsIcon="1"');
        if (c.rId) attrs.push(`r:id="${c.rId}"`);
        if (c.imageWidth !== undefined) attrs.push(`imgW="${c.imageWidth}"`);
        if (c.imageHeight !== undefined) attrs.push(`imgH="${c.imageHeight}"`);
        return `<p:control ${attrs.join(" ")}/>`;
      })
      .join("");
    parts.push(`<p:controls>${ctrlItems}</p:controls>`);
  }

  parts.push("</p:cSld>");

  // p:clrMapOvr
  parts.push("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");

  // p:transition
  if (slideOpts.transition) {
    parts.push(buildTransition(slideOpts.transition));
  }

  // p:timing (animation) — slide level
  if (slideOpts.animations && slideOpts.animations.length > 0) {
    const entries = slideOpts.animations.map((a) => ({ spid: a.shapeId, options: a.options }));
    parts.push(timingDesc.stringify({ entries }, ctx) ?? "");
  }

  parts.push("</p:sld>");
  return parts.join("");
}

// ── Main compiler entry ──

/**
 * Compile a PPTX presentation using the descriptor pipeline.
 *
 * Uses File for structural setup and descriptors for slide serialization.
 */
export function compilePresentation(
  file: File,
  overrides: readonly XmlifyedFile[] = [],
  mediaLevel: number = 0,
): Zippable {
  const descCtx = new PptxWriteContext();

  const mapping: XmlifyedFileMapping = {
    AppProperties: {
      data: XML_DECL + APP_PROPS_XML,
      path: "docProps/app.xml",
    },
    Properties: {
      data: XML_DECL + buildCorePropertiesXmlString(file.corePropsOptions),
      path: "docProps/core.xml",
    },
    FileRelationships: {
      data: XML_DECL + file.fileRelationships.serialize(),
      path: "_rels/.rels",
    },
  };

  // Themes
  const themes = file.themes;
  for (let ti = 0; ti < themes.length; ti++) {
    mapping[`Theme${ti}`] = {
      data: XML_DECL + themes[ti].serialize(),
      path: `ppt/theme/theme${ti + 1}.xml`,
    };
  }

  mapping["TableStyles"] = {
    data: XML_DECL + (tableStylesDesc.stringify({ opts: file.tableStylesOpts }, descCtx) ?? ""),
    path: "ppt/tableStyles.xml",
  };

  mapping["PresProps"] = {
    data: XML_DECL + (presPropsDesc.stringify(file.presPropsFullOpts ?? {}, descCtx) ?? ""),
    path: "ppt/presProps.xml",
  };

  mapping["ViewProps"] = {
    data: XML_DECL + (viewPropsDesc.stringify(file.viewOpts ?? {}, descCtx) ?? ""),
    path: "ppt/viewProps.xml",
  };

  // Slide Masters
  const masters = file.slideMasters;
  const masterRels = file.slideMasterRelsArray;
  for (let mi = 0; mi < masters.length; mi++) {
    mapping[`SlideMaster${mi}`] = {
      data: XML_DECL + (slideMasterDesc.stringify({ master: masters[mi] }, descCtx) ?? ""),
      path: `ppt/slideMasters/slideMaster${mi + 1}.xml`,
    };
    mapping[`SlideMasterRels${mi}`] = {
      data: XML_DECL + masterRels[mi].serialize(),
      path: `ppt/slideMasters/_rels/slideMaster${mi + 1}.xml.rels`,
    };
  }

  // Slide Layouts
  const layouts = file.allLayouts;
  const layoutRels = file.allLayoutRelsArray;
  for (let li = 0; li < layouts.length; li++) {
    mapping[`SlideLayout${li}`] = {
      data: XML_DECL + (slideLayoutDesc.stringify({ layout: layouts[li].layout }, descCtx) ?? ""),
      path: `ppt/slideLayouts/slideLayout${li + 1}.xml`,
    };
    mapping[`SlideLayoutRels${li}`] = {
      data: XML_DECL + layoutRels[li].serialize(),
      path: `ppt/slideLayouts/_rels/slideLayout${li + 1}.xml.rels`,
    };
  }

  // Register content types
  for (let i = 0; i < layouts.length; i++) file.contentTypes.addSlideLayout(i + 1);
  for (let mi = 0; mi < masters.length; mi++) {
    file.contentTypes.addSlideMaster(mi + 1);
    file.contentTypes.addTheme(mi + 1);
  }

  // Notes Master
  if (file.notesTexts.length > 0) {
    const notesMasterRId = file.presRels.relationshipCount + 1;
    file.presRels.addRelationship(
      notesMasterRId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
      "notesMasters/notesMaster1.xml",
    );
    file.presOptions.notesMasterRId = notesMasterRId;
    const notesMasterThemeIndex = themes.length + 1;
    mapping["NotesMaster"] = {
      data:
        XML_DECL + (notesMasterDesc.stringify({ options: file.notesMasterOptions }, descCtx) ?? ""),
      path: "ppt/notesMasters/notesMaster1.xml",
    };
    const notesMasterTheme = new DefaultTheme();
    mapping["NotesMasterTheme"] = {
      data: XML_DECL + notesMasterTheme.serialize(),
      path: `ppt/theme/theme${notesMasterThemeIndex}.xml`,
    };
    file.contentTypes.addTheme(notesMasterThemeIndex);
    const notesMasterRels = new Relationships();
    notesMasterRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      `../theme/theme${notesMasterThemeIndex}.xml`,
    );
    mapping["NotesMasterRelationships"] = {
      data: XML_DECL + notesMasterRels.serialize(),
      path: "ppt/notesMasters/_rels/notesMaster1.xml.rels",
    };
  }

  // Handout Master
  if (file.hasHandoutMaster) {
    const handoutMasterRId = file.presRels.relationshipCount + 1;
    file.presRels.addRelationship(
      handoutMasterRId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster",
      "handoutMasters/handoutMaster1.xml",
    );
    file.presOptions.handoutMasterRId = handoutMasterRId;
    const handoutMasterThemeIndex = themes.length + (file.notesTexts.length > 0 ? 2 : 1);
    mapping["HandoutMaster"] = {
      data:
        XML_DECL +
        (handoutMasterDesc.stringify({ options: file.handoutMasterOptions }, descCtx) ?? ""),
      path: "ppt/handoutMasters/handoutMaster1.xml",
    };
    const handoutMasterTheme = new DefaultTheme();
    mapping["HandoutMasterTheme"] = {
      data: XML_DECL + handoutMasterTheme.serialize(),
      path: `ppt/theme/theme${handoutMasterThemeIndex}.xml`,
    };
    file.contentTypes.addTheme(handoutMasterThemeIndex);
    const handoutMasterRels = new Relationships();
    handoutMasterRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      `../theme/theme${handoutMasterThemeIndex}.xml`,
    );
    mapping["HandoutMasterRelationships"] = {
      data: XML_DECL + handoutMasterRels.serialize(),
      path: "ppt/handoutMasters/_rels/handoutMaster1.xml.rels",
    };
  }

  // Comment Authors
  if (file.commentAuthorEntries) {
    file.presRels.addRelationship(
      file.presRels.relationshipCount + 1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
      "commentAuthors.xml",
    );
  }

  // Presentation XML
  const presBody = presentationDesc.stringify(file.presOptions, descCtx);
  const presentationXml = presBody ? XML_DECL + presBody : "";
  const mediaData = getReferencedMedia(presentationXml, file.media.array);
  const presImageOffset = file.presRels.relationshipCount + 1;
  for (let idx = 0; idx < mediaData.length; idx++) {
    file.presRels.addRelationship(
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

  mapping["Presentation"] = {
    data: replacedPresentationXml,
    path: "ppt/presentation.xml",
  };
  mapping["PresentationRelationships"] = {
    data: XML_DECL + file.presRels.serialize(),
    path: "ppt/_rels/presentation.xml.rels",
  };

  // Slides — use descriptor pipeline for serialization
  for (let i = 0; i < file.slideCount; i++) {
    // Use descriptor-based slide serializer
    const slideXml = stringifySlide(file.slideOptions[i], descCtx);

    const slideMediaData = getReferencedMedia(slideXml, file.media.array);
    const slideRels = file.slideRelationships[i];
    const slideImageOffset = slideRels.relationshipCount + 1;
    for (let idx = 0; idx < slideMediaData.length; idx++) {
      slideRels.addRelationship(
        slideImageOffset + idx,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/${slideMediaData[idx].fileName}`,
      );
    }

    let replacedSlideXml = replaceImagePlaceholders(slideXml, slideMediaData, slideImageOffset);

    if (hasPlaceholders(replacedSlideXml)) {
      // Chart — replace placeholder with rId
      const slideChartKeys = collectPlaceholderKeys(replacedSlideXml, "chart:");
      if (slideChartKeys.length > 0) {
        const slideChartOffset = slideRels.relationshipCount + 1;
        const slideChartKeySet = new Set(slideChartKeys);
        // Charts may come from XmlComponent path (file.charts) or descriptor path (descCtx.charts)
        const xmlCompCharts = file.charts.array.filter((c) => slideChartKeySet.has(c.key));
        const descCharts = descCtx.charts.filter((c) => slideChartKeySet.has(c.key));
        const allChartKeys = [...xmlCompCharts.map((c) => c.key), ...descCharts.map((c) => c.key)];
        replacedSlideXml = replaceChartPlaceholders(
          replacedSlideXml,
          allChartKeys,
          slideChartOffset,
        );
        for (let ci = 0; ci < allChartKeys.length; ci++) {
          slideRels.addRelationship(
            slideChartOffset + ci,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            `../charts/chart${getChartGlobalIndex(allChartKeys[ci], file.charts.array, descCtx.charts) + 1}.xml`,
          );
        }
      }

      // SmartArt — replace placeholders with rIds
      const slideSmartArtKeys = collectPlaceholderKeys(replacedSlideXml, "smartart:");
      if (slideSmartArtKeys.length > 0) {
        const slideSmartArtKeySet = new Set(slideSmartArtKeys);
        const xmlCompSmartArts = file.smartArts.array.filter((s) => slideSmartArtKeySet.has(s.key));
        const descSmartArts = descCtx.smartArts.filter((s) => slideSmartArtKeySet.has(s.key));
        const allSaKeys = [
          ...xmlCompSmartArts.map((s) => s.key),
          ...descSmartArts.map((s) => s.key),
        ];
        const saOffset = slideRels.relationshipCount + 1;
        replacedSlideXml = replaceSmartArtPlaceholders(replacedSlideXml, allSaKeys, saOffset);
        const saGlobalStart = computeSmartArtGlobalStart(
          allSaKeys[0],
          file.smartArts.array,
          descCtx.smartArts,
        );
        addSmartArtRelationships(
          allSaKeys,
          (id, type, target) => {
            slideRels.addRelationship(id, type, target);
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

      // Hyperlinks
      const slideHlinkKeys = collectPlaceholderKeys(replacedSlideXml, "hlink:");
      if (slideHlinkKeys.length > 0) {
        const slideHlinkKeySet = new Set(slideHlinkKeys);
        const slideHlinks = file.hyperlinks.array.filter((h) => slideHlinkKeySet.has(h.key));
        const hlinkOffset = slideRels.relationshipCount + 1;
        replacedSlideXml = replaceHyperlinkPlaceholders(replacedSlideXml, slideHlinks, hlinkOffset);
        for (let hi = 0; hi < slideHlinks.length; hi++) {
          slideRels.addRelationship(
            hlinkOffset + hi,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            slideHlinks[hi].url,
            "External",
          );
        }
      }

      // Media (video/audio)
      const slideMediaRefs = getMediaRefs(replacedSlideXml, file.media.array);
      const slideVideoRefs = getVideoRefs(replacedSlideXml, file.media.array);
      if (slideMediaRefs.length > 0 || slideVideoRefs.length > 0) {
        const mediaOffset = slideRels.relationshipCount + 1;
        const videoOffset = mediaOffset + slideMediaRefs.length;
        replacedSlideXml = replaceMediaPlaceholders(replacedSlideXml, slideMediaRefs, mediaOffset);
        replacedSlideXml = replaceVideoPlaceholders(replacedSlideXml, slideVideoRefs, videoOffset);
        for (let mi = 0; mi < slideMediaRefs.length; mi++) {
          slideRels.addRelationship(
            mediaOffset + mi,
            "http://schemas.microsoft.com/office/2007/relationships/media",
            `../media/${slideMediaRefs[mi].fileName}`,
          );
        }
        for (let vi = 0; vi < slideVideoRefs.length; vi++) {
          slideRels.addRelationship(
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

    // Comment relationship
    const slideCommentEntries = file.slideCommentEntries[i];
    if (slideCommentEntries) {
      slideRels.addRelationship(
        slideRels.relationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        `../comments/comment${i + 1}.xml`,
      );
    }

    // Notes slide relationship
    const notesSlideIndex = file.notesSlideIndexMap.get(i);
    if (notesSlideIndex !== undefined) {
      slideRels.addRelationship(
        slideRels.relationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
        `../notesSlides/notesSlide${notesSlideIndex + 1}.xml`,
      );
    }

    // Sync properties relationship
    const slideSyncIndex = file.slideSyncIndexMap.get(i);
    if (slideSyncIndex !== undefined) {
      slideRels.addRelationship(
        slideRels.relationshipCount + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideSyncProperties",
        `../slideSyncPr/slideSyncPr${slideSyncIndex + 1}.xml`,
      );
    }

    mapping[`SlideRelationships${i}`] = {
      data: XML_DECL + slideRels.serialize(),
      path: `ppt/slides/_rels/slide${i + 1}.xml.rels`,
    };
  }

  // Content Types — charts (from file.charts + descCtx.charts)
  const chartCountFromXmlComponents = file.charts.array.length;
  for (let i = 0; i < chartCountFromXmlComponents; i++) file.contentTypes.addChart(i + 1);
  for (let i = 0; i < descCtx.charts.length; i++)
    file.contentTypes.addChart(chartCountFromXmlComponents + i + 1);

  // Content Types — smartarts (from file.smartArts + descCtx.smartArts)
  const saCountFromXmlComponents = file.smartArts.array.length;
  for (let i = 0; i < saCountFromXmlComponents; i++) {
    file.contentTypes.addDiagramData(i + 1);
    file.contentTypes.addDiagramLayout(i + 1);
    file.contentTypes.addDiagramStyle(i + 1);
    file.contentTypes.addDiagramColors(i + 1);
    file.contentTypes.addDiagramDrawing(i + 1);
  }
  for (let i = 0; i < descCtx.smartArts.length; i++) {
    const idx = saCountFromXmlComponents + i + 1;
    file.contentTypes.addDiagramData(idx);
    file.contentTypes.addDiagramLayout(idx);
    file.contentTypes.addDiagramStyle(idx);
    file.contentTypes.addDiagramColors(idx);
    file.contentTypes.addDiagramDrawing(idx);
  }
  mapping["ContentTypes"] = {
    data: XML_DECL + (contentTypesDesc.stringify({ builder: file.contentTypes }, descCtx) ?? ""),
    path: "[Content_Types].xml",
  };

  // Compile mapping to Zippable
  const files = compileMapping(mapping, overrides);

  // Chart parts (XmlComponent charts serialized via Formatter + descriptor charts as pre-built XML)
  const allCharts = [
    ...file.charts.array.map((c) => ({
      key: c.key,
      xml: XML_DECL + c.chartSpace.serialize(),
    })),
    ...descCtx.charts.map((c) => ({ key: c.key, xml: c.chartSpaceXml })),
  ];
  for (let i = 0; i < allCharts.length; i++) {
    files[`ppt/charts/chart${i + 1}.xml`] = encoder.encode(allCharts[i].xml);
    files[`ppt/charts/_rels/chart${i + 1}.xml.rels`] = encoder.encode(
      `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
    );
  }

  // SmartArt parts (XmlComponent data models + descriptor pre-built XML)
  const allSmartArts = [
    ...file.smartArts.array.map((s) => ({
      key: s.key,
      dataModelXml: XML_DECL + s.dataModel.serialize(),
      layout: s.layout,
      style: s.style,
      color: s.color,
    })),
    ...descCtx.smartArts.map((s) => ({
      key: s.key,
      dataModelXml: s.dataModelXml,
      layout: s.layout,
      style: s.style,
      color: s.color,
    })),
  ];
  for (let i = 0; i < allSmartArts.length; i++) {
    const sa = allSmartArts[i];
    files[`ppt/diagrams/data${i + 1}.xml`] = encoder.encode(sa.dataModelXml);
    files[`ppt/diagrams/layout${i + 1}.xml`] = encoder.encode(getLayoutXml(sa.layout));
    files[`ppt/diagrams/quickStyle${i + 1}.xml`] = encoder.encode(getStyleXml(sa.style));
    files[`ppt/diagrams/colors${i + 1}.xml`] = encoder.encode(getColorXml(sa.color));
    files[`ppt/diagrams/drawing${i + 1}.xml`] = encoder.encode(DEFAULT_DRAWING_XML);
  }

  // ViewProps relationships
  if (file.hasOutlineViewSlides) {
    const vpRels = new Relationships();
    for (let i = 0; i < file.slideCount; i++) {
      vpRels.addRelationship(
        i + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        `slides/slide${i + 1}.xml`,
      );
    }
    files["ppt/_rels/viewProps.xml.rels"] = encoder.encode(XML_DECL + vpRels.serialize());
  }

  // PresProps relationships
  const htmlPubInfo = file.htmlPublishInfo;
  if (htmlPubInfo) {
    const presPropsRels = new Relationships();
    presPropsRels.addRelationship(
      htmlPubInfo.rId.replace("rId", ""),
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      htmlPubInfo.target ?? "presentation.htm",
      "External",
    );
    files["ppt/_rels/presProps.xml.rels"] = encoder.encode(XML_DECL + presPropsRels.serialize());
  }

  // Notes slides
  const notesSlideToSlide = new Map<number, number>();
  for (const [slideIdx, notesIdx] of file.notesSlideIndexMap) {
    notesSlideToSlide.set(notesIdx, slideIdx);
  }
  const notesTexts = file.notesTexts;
  for (let i = 0; i < notesTexts.length; i++) {
    files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = encoder.encode(
      XML_DECL + (notesSlideDesc.stringify({ text: notesTexts[i] }, descCtx) ?? ""),
    );
    const slideIdx = notesSlideToSlide.get(i) ?? 0;
    const nsRels = new Relationships();
    nsRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
      "../notesMasters/notesMaster1.xml",
    );
    nsRels.addRelationship(
      2,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
      `../slides/slide${slideIdx + 1}.xml`,
    );
    files[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = encoder.encode(
      XML_DECL + nsRels.serialize(),
    );
  }

  // Slide sync properties
  const slideSyncOpts = file.slideSyncOptionsList;
  for (let i = 0; i < slideSyncOpts.length; i++) {
    files[`ppt/slideSyncPr/slideSyncPr${i + 1}.xml`] = encoder.encode(
      XML_DECL + (slideSyncDesc.stringify(slideSyncOpts[i], descCtx) ?? ""),
    );
  }

  // Comment authors
  if (file.commentAuthorEntries) {
    files["ppt/commentAuthors.xml"] = encoder.encode(
      XML_DECL + (commentAuthorsDesc.stringify(file.commentAuthorEntries, descCtx) ?? ""),
    );
  }

  // Slide comments
  const commentEntries = file.slideCommentEntries;
  for (let i = 0; i < commentEntries.length; i++) {
    if (commentEntries[i]) {
      files[`ppt/comments/comment${i + 1}.xml`] = encoder.encode(
        XML_DECL + (slideCommentsDesc.stringify(commentEntries[i]!, descCtx) ?? ""),
      );
    }
  }

  // Media files
  for (const image of file.media.array) {
    files[`ppt/media/${image.fileName}`] = [
      image.data,
      { level: mediaLevel as ZipOptions["level"] },
    ];
    if (image.type === "svg" && "fallback" in image) {
      const fallback = (
        image as IMediaData & {
          readonly fallback: { readonly fileName: string; readonly data: Uint8Array };
        }
      ).fallback;
      files[`ppt/media/${fallback.fileName}`] = [
        fallback.data,
        { level: mediaLevel as ZipOptions["level"] },
      ];
    }
  }

  return files;
}

/** Compute the global chart index across legacy + descriptor charts. */
function getChartGlobalIndex(
  key: string,
  legacyCharts: readonly { readonly key: string }[],
  descCharts: readonly { readonly key: string }[],
): number {
  const legacyIdx = legacyCharts.findIndex((c) => c.key === key);
  if (legacyIdx >= 0) return legacyIdx;
  return legacyCharts.length + descCharts.findIndex((c) => c.key === key);
}

/** Compute the global SmartArt start index for the first key on a slide. */
function computeSmartArtGlobalStart(
  firstKey: string,
  legacySmartArts: readonly { readonly key: string }[],
  descSmartArts: readonly { readonly key: string }[],
): number {
  const legacyIdx = legacySmartArts.findIndex((s) => s.key === firstKey);
  if (legacyIdx >= 0) return legacyIdx;
  return legacySmartArts.length + descSmartArts.findIndex((s) => s.key === firstKey);
}
