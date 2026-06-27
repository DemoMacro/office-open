/**
 * Presentation (p:presentation) descriptor for PPTX.
 *
 * @module
 */

import { derivePasswordHash, uniqueUuid } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, escapeXml, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type {
  PresentationPartOptions,
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
} from "@parts/presentation";

// Re-export sub-types for convenience
export type {
  PresentationPartOptions as PresentationDescriptorOptions,
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
};

// ── Default text style ──

const DEFAULT_TEXT_STYLE_XML = `<p:defaultTextStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:defPPr><a:defRPr/></a:defPPr>
  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl1pPr>
  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl2pPr>
  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl3pPr>
  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl4pPr>
  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl5pPr>
  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl6pPr>
  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl7pPr>
  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl8pPr>
  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
    <a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>
  </a:lvl9pPr>
</p:defaultTextStyle>`;

// ── Stringify (internal) ──

function stringifyPresentation(opts: PresentationPartOptions): string {
  const cx = opts.slideWidth ?? 12192000;
  const cy = opts.slideHeight ?? 6858000;

  const rootAttrs: string[] = [];
  if (opts.serverZoom) rootAttrs.push(` serverZoom="${opts.serverZoom}"`);
  if (opts.firstSlideNum !== undefined) rootAttrs.push(` firstSlideNum="${opts.firstSlideNum}"`);
  if (opts.showSpecialPlsOnTitleSld === false) rootAttrs.push(' showSpecialPlsOnTitleSld="0"');
  if (opts.rtl) rootAttrs.push(' rtl="1"');
  if (opts.removePersonalInfoOnSave) rootAttrs.push(' removePersonalInfoOnSave="1"');
  if (opts.compatMode) rootAttrs.push(' compatMode="1"');
  if (opts.strictFirstAndLastChars === false) rootAttrs.push(' strictFirstAndLastChars="0"');
  if (opts.embedTrueTypeFonts) rootAttrs.push(' embedTrueTypeFonts="1"');
  if (opts.saveSubsetFonts) rootAttrs.push(' saveSubsetFonts="1"');
  if (opts.autoCompressPictures === false) rootAttrs.push(' autoCompressPictures="0"');
  if (opts.bookmarkIdSeed !== undefined) rootAttrs.push(` bookmarkIdSeed="${opts.bookmarkIdSeed}"`);
  if (opts.conformance) rootAttrs.push(` conformance="${opts.conformance}"`);

  const parts: string[] = [];
  parts.push(
    `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${rootAttrs.join("")}>`,
  );

  // sldMasterIdLst
  parts.push("<p:sldMasterIdLst>");
  for (let mi = 0; mi < opts.masterCount; mi++) {
    parts.push(`<p:sldMasterId id="${2147483648 + mi * 12}" r:id="rId${mi + 1}"/>`);
  }
  parts.push("</p:sldMasterIdLst>");

  if (opts.notesMasterRId) {
    parts.push(
      `<p:notesMasterIdLst><p:notesMasterId r:id="rId${opts.notesMasterRId}"/></p:notesMasterIdLst>`,
    );
  }
  if (opts.handoutMasterRId) {
    parts.push(
      `<p:handoutMasterIdLst><p:handoutMasterId r:id="rId${opts.handoutMasterRId}"/></p:handoutMasterIdLst>`,
    );
  }

  // sldIdLst
  parts.push("<p:sldIdLst>");
  const sRidOff = opts.masterCount + 1;
  for (let i = 0; i < opts.slideIds.length; i++) {
    parts.push(`<p:sldId id="${opts.slideIds[i]}" r:id="rId${sRidOff + i}"/>`);
  }
  parts.push("</p:sldIdLst>");

  parts.push(`<p:sldSz cx="${cx}" cy="${cy}"/>`);
  parts.push('<p:notesSz cx="6858000" cy="9144000"/>');

  // embeddedFontLst
  if (opts.embeddedFonts && opts.embeddedFonts.length > 0) {
    parts.push("<p:embeddedFontLst>");
    for (const ef of opts.embeddedFonts) {
      const fontAttrs: string[] = [`typeface="${ef.font.typeface}"`];
      if (ef.font.panose) fontAttrs.push(`panose="${ef.font.panose}"`);
      if (ef.font.pitchFamily !== undefined) fontAttrs.push(`pitchFamily="${ef.font.pitchFamily}"`);
      if (ef.font.charset !== undefined) fontAttrs.push(`charset="${ef.font.charset}"`);
      parts.push("<p:embeddedFont>");
      parts.push(`<p:font ${fontAttrs.join(" ")}/>`);
      if (ef.regular) parts.push(`<p:regular r:id="${ef.regular}"/>`);
      if (ef.bold) parts.push(`<p:bold r:id="${ef.bold}"/>`);
      if (ef.italic) parts.push(`<p:italic r:id="${ef.italic}"/>`);
      if (ef.boldItalic) parts.push(`<p:boldItalic r:id="${ef.boldItalic}"/>`);
      parts.push("</p:embeddedFont>");
    }
    parts.push("</p:embeddedFontLst>");
  }

  // custShowLst
  if (opts.customShows && opts.customShows.length > 0) {
    parts.push("<p:custShowLst>");
    for (const cs of opts.customShows) {
      parts.push(`<p:custShow name="${cs.name}" id="${cs.id}">`);
      parts.push("<p:sldLst>");
      for (const sl of cs.slides) {
        parts.push(`<p:sld r:id="${sl.rId}"/>`);
      }
      parts.push("</p:sldLst>");
      parts.push("</p:custShow>");
    }
    parts.push("</p:custShowLst>");
  }

  // photoAlbum
  if (opts.photoAlbum) {
    const pa = opts.photoAlbum;
    const paAttrs: string[] = [];
    if (pa.blackWhite) paAttrs.push(' bw="1"');
    if (pa.showCaptions) paAttrs.push(' showCaptions="1"');
    if (pa.layout) paAttrs.push(` layout="${pa.layout}"`);
    if (pa.frame) paAttrs.push(` frame="${pa.frame}"`);
    parts.push(`<p:photoAlbum${paAttrs.join("")}/>`);
  }

  // custDataLst
  if (opts.customerData) {
    const cdParts: string[] = [];
    if (opts.customerData.data) {
      for (const d of opts.customerData.data) {
        cdParts.push(`<p:custData r:id="${d.rId}"/>`);
      }
    }
    if (opts.customerData.tags) {
      cdParts.push(`<p:tags r:id="${opts.customerData.tags.rId}"/>`);
    }
    if (opts.customerData.tagList) {
      const tagParts = opts.customerData.tagList
        .map((t) => `<p:tag name="${t.name}" val="${t.val}"/>`)
        .join("");
      cdParts.push(`<p:tagLst>${tagParts}</p:tagLst>`);
    }
    if (cdParts.length > 0) {
      parts.push(`<p:custDataLst>${cdParts.join("")}</p:custDataLst>`);
    }
  }

  // kinsoku
  if (opts.kinsoku && opts.kinsoku.length > 0) {
    for (const k of opts.kinsoku) {
      const kAttrs: string[] = [];
      if (k.lang) kAttrs.push(` lang="${k.lang}"`);
      kAttrs.push(` invalStChars="${k.invalStChars}"`);
      kAttrs.push(` invalEndChars="${k.invalEndChars}"`);
      parts.push(`<p:kinsoku${kAttrs.join("")}/>`);
    }
  }

  parts.push(DEFAULT_TEXT_STYLE_XML);

  // modifyVerifier
  if (opts.modifyVerifier) {
    const mv = opts.modifyVerifier;
    let derived: ReturnType<typeof derivePasswordHash> | undefined;
    if (mv.password !== undefined && mv.hashValue === undefined && mv.hashData === undefined) {
      derived = derivePasswordHash(mv.password);
    }
    const useTransitional =
      mv.cryptoProviderType !== undefined || mv.cryptoAlgorithmSid !== undefined;
    const mvAttrs: string[] = [];

    if (useTransitional) {
      const hashData = mv.hashData ?? derived?.hashValue;
      const saltData = mv.saltData ?? derived?.saltValue;
      const spinCount = mv.spinCount ?? derived?.spinCount;
      if (mv.cryptoProviderType) mvAttrs.push(` cryptProviderType="${mv.cryptoProviderType}"`);
      if (mv.cryptoAlgorithmClass)
        mvAttrs.push(` cryptAlgorithmClass="${mv.cryptoAlgorithmClass}"`);
      if (mv.cryptoAlgorithmType) mvAttrs.push(` cryptAlgorithmType="${mv.cryptoAlgorithmType}"`);
      if (mv.cryptoAlgorithmSid !== undefined)
        mvAttrs.push(` cryptAlgorithmSid="${mv.cryptoAlgorithmSid}"`);
      if (spinCount !== undefined) mvAttrs.push(` spinCount="${spinCount}"`);
      if (saltData) mvAttrs.push(` saltData="${saltData}"`);
      if (hashData) mvAttrs.push(` hashData="${hashData}"`);
      if (mv.cryptoProvider) mvAttrs.push(` cryptProvider="${mv.cryptoProvider}"`);
      if (mv.algorithmExtensionId !== undefined)
        mvAttrs.push(` algIdExt="${mv.algorithmExtensionId}"`);
      if (mv.algorithmExtensionSource)
        mvAttrs.push(` algIdExtSource="${mv.algorithmExtensionSource}"`);
      if (mv.cryptoProviderTypeExtension !== undefined)
        mvAttrs.push(` cryptProviderTypeExt="${mv.cryptoProviderTypeExtension}"`);
      if (mv.cryptoProviderTypeExtensionSource)
        mvAttrs.push(` cryptProviderTypeExtSource="${mv.cryptoProviderTypeExtensionSource}"`);
    } else {
      const hashValue = mv.hashValue ?? derived?.hashValue;
      const saltValue = mv.saltValue ?? derived?.saltValue;
      const algorithmName = mv.algorithmName ?? derived?.algorithmName;
      if (algorithmName) mvAttrs.push(` algorithmName="${algorithmName}"`);
      if (hashValue) mvAttrs.push(` hashValue="${hashValue}"`);
      if (saltValue) mvAttrs.push(` saltValue="${saltValue}"`);
      if (mv.spinValue !== undefined) mvAttrs.push(` spinValue="${mv.spinValue}"`);
    }

    parts.push(`<p:modifyVerifier${mvAttrs.join("")}/>`);
  }

  // p14:sectionLst — slide sections (Microsoft PowerPoint 2010 extension).
  // Only emitted when at least one slide carries a section name.
  if (opts.sections && opts.sections.length > 0) {
    const sectionXml: string[] = [];
    for (const sec of opts.sections) {
      const ids = sec.slideIndices
        .map((idx) => opts.slideIds[idx])
        .filter((id): id is number => id !== undefined);
      if (ids.length === 0) continue;
      const sldIds = ids.map((id) => `<p14:sldId id="${id}"/>`).join("");
      sectionXml.push(
        `<p14:section name="${escapeXml(sec.name)}" id="{${uniqueUuid().toUpperCase()}}"><p14:sldIdLst>${sldIds}</p14:sldIdLst></p14:section>`,
      );
    }
    if (sectionXml.length > 0) {
      parts.push(
        `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${sectionXml.join("")}</p14:sectionLst></p:ext></p:extLst>`,
      );
    }
  }

  parts.push("</p:presentation>");
  return parts.join("");
}

// ── Parse ──

function parsePresentation(el: XmlElement): PresentationPartOptions {
  const result: Partial<PresentationPartOptions> = {};

  if (el.attributes) {
    const a = el.attributes;
    if (a["serverZoom"]) result.serverZoom = String(a["serverZoom"]);
    if (a["firstSlideNum"] !== undefined) result.firstSlideNum = Number(a["firstSlideNum"]);
    if (a["showSpecialPlsOnTitleSld"] !== undefined)
      result.showSpecialPlsOnTitleSld = a["showSpecialPlsOnTitleSld"] !== "0";
    if (a["rtl"] !== undefined) result.rtl = a["rtl"] === "1";
    if (a["removePersonalInfoOnSave"] !== undefined)
      result.removePersonalInfoOnSave = a["removePersonalInfoOnSave"] === "1";
    if (a["compatMode"] !== undefined) result.compatMode = a["compatMode"] === "1";
    if (a["strictFirstAndLastChars"] !== undefined)
      result.strictFirstAndLastChars = a["strictFirstAndLastChars"] !== "0";
    if (a["embedTrueTypeFonts"] !== undefined)
      result.embedTrueTypeFonts = a["embedTrueTypeFonts"] === "1";
    if (a["saveSubsetFonts"] !== undefined) result.saveSubsetFonts = a["saveSubsetFonts"] === "1";
    if (a["autoCompressPictures"] !== undefined)
      result.autoCompressPictures = a["autoCompressPictures"] !== "0";
    if (a["bookmarkIdSeed"] !== undefined) result.bookmarkIdSeed = Number(a["bookmarkIdSeed"]);
    if (a["conformance"])
      result.conformance = a["conformance"] as PresentationPartOptions["conformance"];
  }

  // sldSz
  const sldSz = findChild(el, "p:sldSz");
  if (sldSz?.attributes) {
    if (sldSz.attributes["cx"] !== undefined) result.slideWidth = Number(sldSz.attributes["cx"]);
    if (sldSz.attributes["cy"] !== undefined) result.slideHeight = Number(sldSz.attributes["cy"]);
  }

  // sldIdLst
  const sldIdLst = findChild(el, "p:sldIdLst");
  if (sldIdLst) {
    const slideIds: number[] = [];
    for (const child of sldIdLst.elements ?? []) {
      if (child.name === "p:sldId") {
        const id = attrNum(child, "id");
        if (id !== undefined) slideIds.push(id);
      }
    }
    if (slideIds.length > 0) result.slideIds = slideIds;
  }

  // sldMasterIdLst — count masters
  const sldMasterIdLst = findChild(el, "p:sldMasterIdLst");
  if (sldMasterIdLst) {
    let masterCount = 0;
    for (const child of sldMasterIdLst.elements ?? []) {
      if (child.name === "p:sldMasterId") masterCount++;
    }
    result.masterCount = masterCount;
  }

  // notesMasterIdLst
  const notesMasterIdLst = findChild(el, "p:notesMasterIdLst");
  if (notesMasterIdLst) {
    const nmId = findChild(notesMasterIdLst, "p:notesMasterId");
    if (nmId) {
      const rId = attr(nmId, "r:id");
      if (rId) result.notesMasterRId = Number(rId.replace("rId", ""));
    }
  }

  // handoutMasterIdLst
  const handoutMasterIdLst = findChild(el, "p:handoutMasterIdLst");
  if (handoutMasterIdLst) {
    const hmId = findChild(handoutMasterIdLst, "p:handoutMasterId");
    if (hmId) {
      const rId = attr(hmId, "r:id");
      if (rId) result.handoutMasterRId = Number(rId.replace("rId", ""));
    }
  }

  // photoAlbum
  const photoAlbum = findChild(el, "p:photoAlbum");
  if (photoAlbum?.attributes) {
    const pa: PhotoAlbumOptions = {};
    if (photoAlbum.attributes["bw"] === "1") pa.blackWhite = true;
    if (photoAlbum.attributes["showCaptions"] === "1") pa.showCaptions = true;
    if (photoAlbum.attributes["layout"])
      pa.layout = photoAlbum.attributes["layout"] as PhotoAlbumOptions["layout"];
    if (photoAlbum.attributes["frame"])
      pa.frame = photoAlbum.attributes["frame"] as PhotoAlbumOptions["frame"];
    result.photoAlbum = pa;
  }

  // modifyVerifier
  const mv = findChild(el, "p:modifyVerifier");
  if (mv?.attributes) {
    const mvOpts: ModifyVerifierOptions = {};
    if (mv.attributes["algorithmName"])
      mvOpts.algorithmName = String(mv.attributes["algorithmName"]);
    if (mv.attributes["hashValue"]) mvOpts.hashValue = String(mv.attributes["hashValue"]);
    if (mv.attributes["saltValue"]) mvOpts.saltValue = String(mv.attributes["saltValue"]);
    if (mv.attributes["spinValue"] !== undefined)
      mvOpts.spinValue = Number(mv.attributes["spinValue"]);
    if (mv.attributes["cryptProviderType"])
      mvOpts.cryptoProviderType = String(mv.attributes["cryptProviderType"]);
    if (mv.attributes["cryptAlgorithmClass"])
      mvOpts.cryptoAlgorithmClass = String(mv.attributes["cryptAlgorithmClass"]);
    if (mv.attributes["cryptAlgorithmType"])
      mvOpts.cryptoAlgorithmType = String(mv.attributes["cryptAlgorithmType"]);
    if (mv.attributes["cryptAlgorithmSid"] !== undefined)
      mvOpts.cryptoAlgorithmSid = Number(mv.attributes["cryptAlgorithmSid"]);
    if (mv.attributes["spinCount"] !== undefined)
      mvOpts.spinCount = Number(mv.attributes["spinCount"]);
    if (mv.attributes["saltData"]) mvOpts.saltData = String(mv.attributes["saltData"]);
    if (mv.attributes["hashData"]) mvOpts.hashData = String(mv.attributes["hashData"]);
    if (mv.attributes["cryptProvider"])
      mvOpts.cryptoProvider = String(mv.attributes["cryptProvider"]);
    result.modifyVerifier = mvOpts;
  }

  // embeddedFontLst
  const efl = findChild(el, "p:embeddedFontLst");
  if (efl) {
    const fonts: EmbeddedFontOptions[] = [];
    for (const ef of efl.elements ?? []) {
      if (ef.name !== "p:embeddedFont") continue;
      const fontEl = findChild(ef, "p:font");
      if (!fontEl) continue;
      const font: EmbeddedFontOptions["font"] = { typeface: attr(fontEl, "typeface") ?? "" };
      if (fontEl.attributes?.["panose"]) font.panose = String(fontEl.attributes["panose"]);
      const pp = attrNum(fontEl, "pitchFamily");
      if (pp !== undefined) font.pitchFamily = pp;
      const cs = attrNum(fontEl, "charset");
      if (cs !== undefined) font.charset = cs;
      const entry: EmbeddedFontOptions = { font };
      const regular = findChild(ef, "p:regular");
      if (regular) entry.regular = attr(regular, "r:id") ?? "";
      const bold = findChild(ef, "p:bold");
      if (bold) entry.bold = attr(bold, "r:id") ?? "";
      const italic = findChild(ef, "p:italic");
      if (italic) entry.italic = attr(italic, "r:id") ?? "";
      const boldItalic = findChild(ef, "p:boldItalic");
      if (boldItalic) entry.boldItalic = attr(boldItalic, "r:id") ?? "";
      fonts.push(entry);
    }
    if (fonts.length > 0) result.embeddedFonts = fonts;
  }

  // customShows
  const csl = findChild(el, "p:custShowLst");
  if (csl) {
    const shows: CustomShowOptions[] = [];
    for (const cs of csl.elements ?? []) {
      if (cs.name !== "p:custShow") continue;
      const sldLst = findChild(cs, "p:sldLst");
      const slides: { rId: string }[] = [];
      if (sldLst) {
        for (const s of sldLst.elements ?? []) {
          const rId = attr(s, "r:id");
          if (rId) slides.push({ rId });
        }
      }
      shows.push({
        name: attr(cs, "name") ?? "",
        id: attrNum(cs, "id") ?? 0,
        slides,
      });
    }
    if (shows.length > 0) result.customShows = shows;
  }

  // kinsoku
  const kinsokuEls = el.elements?.filter((e) => e.name === "p:kinsoku") ?? [];
  if (kinsokuEls.length > 0) {
    const kinsoku: KinsokuOptions[] = [];
    for (const k of kinsokuEls) {
      const entry: KinsokuOptions = {
        invalStChars: attr(k, "invalStChars") ?? "",
        invalEndChars: attr(k, "invalEndChars") ?? "",
      };
      if (k.attributes?.["lang"]) entry.lang = String(k.attributes["lang"]);
      kinsoku.push(entry);
    }
    result.kinsoku = kinsoku;
  }

  return result as PresentationPartOptions;
}

// ── Descriptor ──

export const presentationDesc: CustomDescriptor<PresentationPartOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyPresentation(opts);
  },

  parse(el, _ctx) {
    return parsePresentation(el);
  },
};
