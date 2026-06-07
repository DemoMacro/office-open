import type { LayoutDefinition } from "@file/file";
import { buildMasterChildrenXml } from "@file/slide/coerce";
import { convertPositionToEmu } from "@office-open/core";

export type SlideLayoutType =
  | "blank"
  | "title"
  | "tx"
  | "twoColTx"
  | "titleOnly"
  | "obj"
  | "secHead"
  | "chart"
  | "tbl"
  | "picTx"
  | "twoObj"
  | "twoTxTwoObj"
  | "objTx"
  | "vertTx"
  | "vertTitleAndTx";

const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;

import { SP_TREE_HEADER } from "@file/constants";

// 16:9 reference slide width
const SW_REF = 12192000;
const sx = (refX: number, sw: number) => Math.round((refX * sw) / SW_REF);

// ── Placeholder XML builders ──

// No xfrm — inherited from master
function dtPlaceholder(id: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Date Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{5BCAD085-E8A6-8845-BD4E-CB4CCA059FC4}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>1/27/13</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

function ftrPlaceholder(id: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Footer Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

function sldNumPlaceholder(id: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Slide Number Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{C1FF6DA9-008F-8B48-92A6-B652298478BF}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// title — no idx, no xfrm (inherited)
function titlePlaceholder(id: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// Body (generic) idx, no xfrm (inherited)
function bodyPlaceholder(id: number, idx: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Content Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// Body with xfrm (positioned)
function bodyPlaceholderAt(
  id: number,
  idx: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Content Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// Body with explicit type="body" and xfrm
function typedBodyPlaceholder(
  id: number,
  idx: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// Title with xfrm
function titlePlaceholderAt(id: number, x: number, y: number, cx: number, cy: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// ctrTitle — centered title for title layout
function ctrTitlePlaceholder(id: number, x: number, y: number, cx: number, cy: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="b"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// subTitle for title layout
function subTitlePlaceholder(id: number, x: number, y: number, cx: number, cy: number): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// Pic placeholder
function picPlaceholder(
  id: number,
  idx: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Picture Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

// ── Layout definitions (16:9 reference positions) ──

interface LayoutDef {
  type: SlideLayoutType;
  name: string;
  buildShapes: (sw: number) => string;
}

const LAYOUT_DEFS: Record<SlideLayoutType, LayoutDef> = {
  // 1. title — ctrTitle + subTitle
  title: {
    type: "title",
    name: "Title Slide",
    buildShapes: (sw) =>
      ctrTitlePlaceholder(2, sx(1524000, sw), 1122363, sx(9144000, sw), 2387600) +
      subTitlePlaceholder(3, sx(1524000, sw), 3602038, sx(9144000, sw), 1655762),
  },
  // 2. obj — title + body (generic content)
  obj: {
    type: "obj",
    name: "Title and Content",
    buildShapes: () => titlePlaceholder(2) + bodyPlaceholder(3, 1),
  },
  // 3. secHead — title + body (both positioned)
  secHead: {
    type: "secHead",
    name: "Section Header",
    buildShapes: (sw) =>
      titlePlaceholderAt(2, sx(831850, sw), 1709738, sx(10515600, sw), 2852737) +
      typedBodyPlaceholder(3, 1, sx(831850, sw), 4589463, sx(10515600, sw), 1500187),
  },
  // 4. twoObj — title + 2 body columns
  twoObj: {
    type: "twoObj",
    name: "Two Content",
    buildShapes: (sw) =>
      titlePlaceholder(2) +
      bodyPlaceholderAt(3, 1, sx(838200, sw), 1825625, sx(5181600, sw), 4351338) +
      bodyPlaceholderAt(4, 2, sx(6172200, sw), 1825625, sx(5181600, sw), 4351338),
  },
  // 5. twoTxTwoObj — title + 4 quadrants (2 text, 2 content)
  twoTxTwoObj: {
    type: "twoTxTwoObj",
    name: "Comparison",
    buildShapes: (sw) =>
      titlePlaceholderAt(2, sx(839788, sw), 365125, sx(10515600, sw), 1325563) +
      typedBodyPlaceholder(3, 1, sx(839788, sw), 1681163, sx(5157787, sw), 823912) +
      bodyPlaceholderAt(4, 2, sx(839788, sw), 2505075, sx(5157787, sw), 3684588) +
      typedBodyPlaceholder(5, 3, sx(6172200, sw), 1681163, sx(5183188, sw), 823912) +
      bodyPlaceholderAt(6, 4, sx(6172200, sw), 2505075, sx(5183188, sw), 3684588),
  },
  // 6. titleOnly — title only
  titleOnly: {
    type: "titleOnly",
    name: "Title Only",
    buildShapes: () => titlePlaceholder(2),
  },
  // 7. blank — no content placeholders
  blank: {
    type: "blank",
    name: "Blank",
    buildShapes: () => "",
  },
  // 8. objTx — title + body + text (object and text)
  objTx: {
    type: "objTx",
    name: "Content with Caption",
    buildShapes: (sw) =>
      titlePlaceholderAt(2, sx(839788, sw), 457200, sx(3932237, sw), 1600200) +
      bodyPlaceholderAt(3, 1, sx(5183188, sw), 987425, sx(6172200, sw), 4873625) +
      typedBodyPlaceholder(4, 2, sx(839788, sw), 2057400, sx(3932237, sw), 3811588),
  },
  // 9. picTx — title + pic + text
  picTx: {
    type: "picTx",
    name: "Picture with Caption",
    buildShapes: (sw) =>
      titlePlaceholderAt(2, sx(839788, sw), 457200, sx(3932237, sw), 1600200) +
      picPlaceholder(3, 1, sx(5183188, sw), 987425, sx(6172200, sw), 4873625) +
      typedBodyPlaceholder(4, 2, sx(839788, sw), 2057400, sx(3932237, sw), 3811588),
  },
  // 10. vertTx — title + body (vertical text)
  vertTx: {
    type: "vertTx",
    name: "Vertical Text",
    buildShapes: () =>
      titlePlaceholder(2) +
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`,
  },
  // 11. vertTitleAndTx — vertical title + body
  vertTitleAndTx: {
    type: "vertTitleAndTx",
    name: "Vertical Title and Text",
    buildShapes: (sw) =>
      titlePlaceholderAt(2, sx(8724900, sw), 365125, sx(2628900, sw), 5811838) +
      typedBodyPlaceholder(3, 1, sx(838200, sw), 365125, sx(7734300, sw), 5811838),
  },
  // Aliases / simplified versions
  tx: {
    type: "tx",
    name: "Title and Text",
    buildShapes: () => titlePlaceholder(2) + bodyPlaceholder(3, 1),
  },
  chart: {
    type: "chart",
    name: "Content with Caption",
    buildShapes: () => titlePlaceholder(2) + bodyPlaceholder(3, 1),
  },
  tbl: {
    type: "tbl",
    name: "Content with Caption",
    buildShapes: () => titlePlaceholder(2) + bodyPlaceholder(3, 1),
  },
  twoColTx: {
    type: "twoColTx",
    name: "Two Content",
    buildShapes: (sw) =>
      titlePlaceholder(2) +
      bodyPlaceholderAt(3, 1, sx(838200, sw), 1825625, sx(5181600, sw), 4351338) +
      bodyPlaceholderAt(4, 2, sx(6172200, sw), 1825625, sx(5181600, sw), 4351338),
  },
};

export function buildLayoutXml(layoutType: SlideLayoutType, slideWidth: number = SW_REF): string {
  const def = LAYOUT_DEFS[layoutType];
  const contentShapes = def.buildShapes(slideWidth);
  // Count content shapes to determine starting ID for footer placeholders
  const contentCount = contentShapes ? (contentShapes.match(/<p:sp>/g) || []).length : 0;
  const footerStartId = 2 + contentCount;

  const footerShapes =
    dtPlaceholder(footerStartId) +
    ftrPlaceholder(footerStartId + 1) +
    sldNumPlaceholder(footerStartId + 2);

  return `<p:sldLayout ${NS} type="${layoutType}" preserve="1"><p:cSld name="${def.name}"><p:spTree>${SP_TREE_HEADER}${contentShapes}${footerShapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

// ── Custom layout builder ──

function positionedTitlePlaceholder(
  id: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

function positionedSubtitlePlaceholder(
  id: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

function positionedBodyPlaceholder(
  id: number,
  idx: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Content Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

function positionedDatePlaceholder(
  id: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Date Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{5BCAD085-E8A6-8845-BD4E-CB4CCA059FC4}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>1/27/13</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

function positionedFooterPlaceholder(
  id: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Footer Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

function positionedSldNumPlaceholder(
  id: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Slide Number Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{C1FF6DA9-008F-8B48-92A6-B652298478BF}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}

export function buildCustomLayoutXml(def: LayoutDefinition): string {
  const ph = def.placeholders ?? {};
  const layoutType = def.type ?? "blank";
  const displayName = def.name ?? LAYOUT_DEFS[layoutType]?.name ?? layoutType;

  const shapes: string[] = [];
  let nextId = 2;

  // Content placeholders
  if (ph.title !== false) {
    const titlePos = ph.title ? convertPositionToEmu(ph.title) : null;
    if (titlePos) {
      shapes.push(
        positionedTitlePlaceholder(nextId++, titlePos.x, titlePos.y, titlePos.cx, titlePos.cy),
      );
    } else {
      shapes.push(titlePlaceholder(nextId++));
    }
  }

  if (ph.subtitle !== false && ph.subtitle !== undefined) {
    const subPos = convertPositionToEmu(ph.subtitle);
    shapes.push(positionedSubtitlePlaceholder(nextId++, subPos.x, subPos.y, subPos.cx, subPos.cy));
  }

  if (ph.body !== false && ph.body !== undefined) {
    const bodyPos = ph.body ? convertPositionToEmu(ph.body) : null;
    if (bodyPos) {
      shapes.push(
        positionedBodyPlaceholder(nextId++, 1, bodyPos.x, bodyPos.y, bodyPos.cx, bodyPos.cy),
      );
    } else {
      shapes.push(bodyPlaceholder(nextId++, 1));
    }
  }

  // Footer placeholders
  if (ph.date !== false && ph.date !== undefined) {
    const datePos = ph.date ? convertPositionToEmu(ph.date) : null;
    if (datePos) {
      shapes.push(
        positionedDatePlaceholder(nextId++, datePos.x, datePos.y, datePos.cx, datePos.cy),
      );
    } else {
      shapes.push(dtPlaceholder(nextId++));
    }
  }

  if (ph.footer !== false && ph.footer !== undefined) {
    const ftrPos = ph.footer ? convertPositionToEmu(ph.footer) : null;
    if (ftrPos) {
      shapes.push(positionedFooterPlaceholder(nextId++, ftrPos.x, ftrPos.y, ftrPos.cx, ftrPos.cy));
    } else {
      shapes.push(ftrPlaceholder(nextId++));
    }
  }

  if (ph.slideNumber !== false && ph.slideNumber !== undefined) {
    const sldPos = ph.slideNumber ? convertPositionToEmu(ph.slideNumber) : null;
    if (sldPos) {
      shapes.push(positionedSldNumPlaceholder(nextId++, sldPos.x, sldPos.y, sldPos.cx, sldPos.cy));
    } else {
      shapes.push(sldNumPlaceholder(nextId++));
    }
  }

  // Children shapes — shift ids
  const childrenXml = buildMasterChildrenXml(def.children);
  if (childrenXml) {
    const offset = nextId - 1;
    shapes.push(childrenXml.replace(/ id="(\d+)"/g, (_, n) => ` id="${parseInt(n) + offset}"`));
  }

  const matchingAttr = def.matchingName !== undefined ? ` matchingName="${def.matchingName}"` : "";
  return `<p:sldLayout ${NS} type="${layoutType}" preserve="1"${matchingAttr}><p:cSld name="${displayName}"><p:spTree>${SP_TREE_HEADER}${shapes.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}
