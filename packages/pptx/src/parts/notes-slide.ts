/**
 * Notes Slide XML builder.
 *
 * Pure function — no XmlComponent inheritance.
 *
 * @module
 */
import { escapeXml } from "@office-open/xml";

export interface NotesSlideOptions {
  text?: string;
}

/**
 * Build a complete p:notes XML string.
 *
 * Replaces the former NotesSlide XmlComponent class hierarchy
 * (NotesSlide → NotesCommonSlideData → NotesShapeTree → SlideImagePlaceholder / NotesBodyPlaceholder).
 */
export function buildNotesSlideXml(options: NotesSlideOptions = {}): string {
  const escapedText = options.text ? escapeXml(options.text) : "";

  return (
    '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    "<p:cSld>" +
    "<p:spTree>" +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    // Slide Image Placeholder
    "<p:sp>" +
    '<p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="685800" y="1600200"/><a:ext cx="5486400" cy="3086100"/></a:xfrm></p:spPr>' +
    "</p:sp>" +
    // Notes Body Placeholder
    "<p:sp>" +
    '<p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="685800" y="4800600"/><a:ext cx="5486400" cy="3600450"/></a:xfrm></p:spPr>' +
    "<p:txBody>" +
    "<a:bodyPr/>" +
    "<a:lstStyle/>" +
    (escapedText
      ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapedText}</a:t></a:r></a:p>`
      : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>') +
    "</p:txBody>" +
    "</p:sp>" +
    "</p:spTree>" +
    "</p:cSld>" +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" +
    "</p:notes>"
  );
}
