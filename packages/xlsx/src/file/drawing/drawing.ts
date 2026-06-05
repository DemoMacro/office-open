/**
 * XLSX Drawing component — generates xl/drawings/drawing{n}.xml.
 *
 * Uses the spreadsheetDrawing namespace (default, no prefix) for anchoring
 * images and charts to worksheet cells.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

export interface ImageOptions {
  /** 1-based column */
  readonly col: number;
  /** Column offset in EMU (default 0) */
  readonly colOffset?: number;
  /** 1-based row */
  readonly row: number;
  /** Row offset in EMU (default 0) */
  readonly rowOffset?: number;
  /** Relationship ID for the image */
  readonly rId: string;
  /** Lock anchor with sheet (default true) */
  readonly locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  readonly printsWithSheet?: boolean;
}

export interface ChartAnchorOptions {
  /** 1-based column */
  readonly col: number;
  /** Column offset in EMU (default 0) */
  readonly colOffset?: number;
  /** 1-based row */
  readonly row: number;
  /** Row offset in EMU (default 0) */
  readonly rowOffset?: number;
  /** Relationship ID for the chart */
  readonly rId: string;
  /** Lock anchor with sheet (default true) */
  readonly locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  readonly printsWithSheet?: boolean;
}

const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const C_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

export class Drawing extends BaseXmlComponent {
  private readonly images: readonly ImageOptions[];
  private readonly charts: readonly ChartAnchorOptions[];

  public constructor(images: readonly ImageOptions[], charts: readonly ChartAnchorOptions[] = []) {
    super("wsDr");
    this.images = images;
    this.charts = charts;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [`<wsDr xmlns="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`];
    let id = 1;
    for (const img of this.images) {
      p.push(
        `<twoCellAnchor editAs="oneCell"><from><col>${img.col - 1}</col><colOff>${img.colOffset ?? 0}</colOff><row>${img.row - 1}</row><rowOff>${img.rowOffset ?? 0}</rowOff></from>`,
        `<to><col>${img.col}</col><colOff>0</colOff><row>${img.row}</row><rowOff>0</rowOff></to>`,
        `<pic><nvPicPr><cNvPr id="${id}" name="Picture ${id}"/><cNvPicPr preferRelativeResize="1"/></nvPicPr>`,
        `<blipFill><a:blip r:embed="${img.rId}"/><a:stretch><a:fillRect/></a:stretch></blipFill>`,
        `<spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></spPr></pic>`,
        `<clientData fLocksWithSheet="${img.locksWithSheet !== false ? 1 : 0}" fPrintsWithSheet="${img.printsWithSheet !== false ? 1 : 0}"/></twoCellAnchor>`,
      );
      id++;
    }
    for (const chart of this.charts) {
      p.push(
        `<twoCellAnchor editAs="oneCell"><from><col>${chart.col - 1}</col><colOff>${chart.colOffset ?? 0}</colOff><row>${chart.row - 1}</row><rowOff>${chart.rowOffset ?? 0}</rowOff></from>`,
        `<to><col>${chart.col + 8}</col><colOff>0</colOff><row>${chart.row + 15}</row><rowOff>0</rowOff></to>`,
        `<graphicFrame><nvGraphicFramePr><cNvPr id="${id}" name="Chart ${id}"/><cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></cNvGraphicFramePr></nvGraphicFramePr>`,
        `<xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xfrm>`,
        `<a:graphic><a:graphicData uri="${C_URI}"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}" r:id="${chart.rId}"/></a:graphicData></a:graphic></graphicFrame>`,
        `<clientData fLocksWithSheet="${chart.locksWithSheet !== false ? 1 : 0}" fPrintsWithSheet="${chart.printsWithSheet !== false ? 1 : 0}"/></twoCellAnchor>`,
      );
      id++;
    }
    p.push("</wsDr>");
    return p.join("");
  }
}
