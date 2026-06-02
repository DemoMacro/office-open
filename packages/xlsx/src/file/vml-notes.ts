/**
 * VML Notes — generates legacy VML drawing for worksheet comments.
 *
 * Excel requires a VML drawing file (xl/drawings/vmlDrawing{N}.vml)
 * referenced by <legacyDrawing r:id="..."/> in the worksheet XML.
 *
 * @module
 */
import type { CommentOptions } from "@file/worksheet";

export class VmlNotes {
  public constructor(private readonly comments: readonly CommentOptions[]) {}

  public toXml(): string {
    const p: string[] = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
      '<v:stroke joinstyle="miter"/>',
      '<v:path gradientshapeok="t" o:connecttype="rect"/>',
      "</v:shapetype>",
    ];

    for (let i = 0; i < this.comments.length; i++) {
      const c = this.comments[i];
      const col = c.cell.charCodeAt(0) - 65;
      const row = parseInt(c.cell.slice(1), 10) - 1;
      // 8-value anchor: fromCol, fromColOffset, fromRow, fromRowOffset, toCol, toColOffset, toRow, toRowOffset
      const anchor = `${col}, 0, ${row}, 0, ${col + 2}, 0, ${row + 2}, 0`;
      p.push(
        `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" ` +
          `style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;` +
          `z-index:1;visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">`,
        `<v:fill color2="infoBackground [80]"/>`,
        `<v:shadow color="none [81]" obscured="t"/>`,
        `<v:path o:connecttype="none"/>`,
        `<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>`,
        `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>`,
        `<x:Anchor>${anchor}</x:Anchor>`,
        `<x:AutoFill>False</x:AutoFill>`,
        `<x:Row>${row}</x:Row>`,
        `<x:Column>${col}</x:Column>`,
        `</x:ClientData>`,
        `</v:shape>`,
      );
    }

    p.push("</xml>");
    return p.join("");
  }
}
