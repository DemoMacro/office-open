import { XmlComponent } from "@file/xml-components";

import { ContinuationSeperatorRun } from "../footnotes/footnote/run/continuation-seperator-run";
import { SeperatorRun } from "../footnotes/footnote/run/seperator-run";
import { LineRuleType, Paragraph } from "../paragraph";
import { Endnote, EndnoteType } from "./endnote/endnote";

export class Endnotes extends XmlComponent {
  public constructor() {
    super("w:endnotes");

    this.root.push({
      _attr: {
        "mc:Ignorable": "w14 w15 wp14",
        "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "xmlns:o": "urn:schemas-microsoft-com:office:office",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:v": "urn:schemas-microsoft-com:vml",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w10": "urn:schemas-microsoft-com:office:word",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
        "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "xmlns:wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "xmlns:wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "xmlns:wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
      },
    });

    const begin = new Endnote({
      children: [
        new Paragraph({
          children: [new SeperatorRun()],
          spacing: {
            after: 0,
            line: 240,
            lineRule: LineRuleType.AUTO,
          },
        }),
      ],
      id: -1,
      type: EndnoteType.SEPARATOR,
    });

    this.root.push(begin);

    const spacing = new Endnote({
      children: [
        new Paragraph({
          children: [new ContinuationSeperatorRun()],
          spacing: {
            after: 0,
            line: 240,
            lineRule: LineRuleType.AUTO,
          },
        }),
      ],
      id: 0,
      type: EndnoteType.CONTINUATION_SEPARATOR,
    });

    this.root.push(spacing);
  }

  public createEndnote(id: number, paragraph: readonly Paragraph[]): void {
    const endnote = new Endnote({
      children: paragraph,
      id: id,
    });

    this.root.push(endnote);
  }
}
