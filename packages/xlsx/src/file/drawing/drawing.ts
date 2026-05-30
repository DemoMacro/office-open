/**
 * XLSX Drawing component — generates xl/drawings/drawing{n}.xml.
 *
 * Uses the spreadsheetDrawing namespace (default, no prefix) for anchoring
 * images and charts to worksheet cells.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

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

  public override prepForXml(_context: Context): IXmlableObject {
    const children: IXmlableObject[] = [
      {
        _attr: {
          xmlns: XDR_NS,
          "xmlns:a": A_NS,
          "xmlns:r": R_NS,
        },
      },
    ];

    let nextId = 1;
    for (const img of this.images) {
      children.push(this.buildImageAnchor(img, nextId++));
    }

    for (const chart of this.charts) {
      children.push(this.buildChartAnchor(chart, nextId++));
    }

    return { wsDr: children };
  }

  private buildFromAnchor(
    col: number,
    row: number,
    colOffset?: number,
    rowOffset?: number,
  ): IXmlableObject {
    return {
      from: [
        { col: [col - 1] },
        { colOff: [colOffset ?? 0] },
        { row: [row - 1] },
        { rowOff: [rowOffset ?? 0] },
      ],
    };
  }

  private buildToAnchor(col: number, row: number): IXmlableObject {
    return {
      to: [{ col: [col] }, { colOff: [0] }, { row: [row] }, { rowOff: [0] }],
    };
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
        `<clientData/></twoCellAnchor>`,
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
        `<clientData/></twoCellAnchor>`,
      );
      id++;
    }
    p.push("</wsDr>");
    return p.join("");
  }

  private buildImageAnchor(img: ImageOptions, id: number): IXmlableObject {
    return {
      twoCellAnchor: [
        { _attr: { editAs: "oneCell" } },
        this.buildFromAnchor(img.col, img.row, img.colOffset, img.rowOffset),
        this.buildToAnchor(img.col, img.row),
        {
          pic: [
            {
              nvPicPr: [
                { cNvPr: { _attr: { id, name: `Picture ${id}` } } },
                { cNvPicPr: [{ _attr: { preferRelativeResize: 1 } }] },
              ],
            },
            {
              blipFill: [
                { "a:blip": { _attr: { "r:embed": img.rId } } },
                { "a:stretch": [{ "a:fillRect": [] }] },
              ],
            },
            {
              spPr: [
                {
                  "a:xfrm": [
                    { "a:off": { _attr: { x: 0, y: 0 } } },
                    { "a:ext": { _attr: { cx: 400000, cy: 300000 } } },
                  ],
                },
                { "a:prstGeom": [{ _attr: { prst: "rect" } }, { "a:avLst": [] }] },
              ],
            },
          ],
        },
        { clientData: [] },
      ],
    };
  }

  private buildChartAnchor(chart: ChartAnchorOptions, id: number): IXmlableObject {
    return {
      twoCellAnchor: [
        { _attr: { editAs: "oneCell" } },
        this.buildFromAnchor(chart.col, chart.row, chart.colOffset, chart.rowOffset),
        this.buildToAnchor(chart.col + 8, chart.row + 15),
        {
          graphicFrame: [
            {
              nvGraphicFramePr: [
                { cNvPr: { _attr: { id, name: `Chart ${id}` } } },
                { cNvGraphicFramePr: [{ "a:graphicFrameLocks": { _attr: { noGrp: 1 } } }] },
              ],
            },
            {
              xfrm: [
                { "a:off": { _attr: { x: 0, y: 0 } } },
                { "a:ext": { _attr: { cx: 0, cy: 0 } } },
              ],
            },
            {
              "a:graphic": [
                {
                  "a:graphicData": [
                    { _attr: { uri: C_URI } },
                    {
                      "c:chart": {
                        _attr: {
                          "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                          "xmlns:r": R_NS,
                          "r:id": chart.rId,
                        },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
        { clientData: [] },
      ],
    };
  }
}
