import { Formatter } from "@export/formatter";
import { BorderStyle } from "@file/border";
import { HorizontalPositionAlign, VerticalPositionAlign } from "@file/shared";
import { EMPTY_OBJECT } from "@file/xml-components";
import { toElement } from "@office-open/xml";
import * as convenienceFunctions from "@util/convenience-functions";
import { afterEach, assert, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { ParseContext } from "../../parse/context";
import { DocumentWrapper } from "../document-wrapper";
import type { ViewWrapper } from "../document-wrapper";
import type { File } from "../file";
import { ShadingType } from "../shading";
import {
  AlignmentType,
  HeadingLevel,
  LeaderType,
  PageBreak,
  TabStopPosition,
  TabStopType,
} from "./formatting";
import { FrameAnchorType } from "./frame";
import { Bookmark, ExternalHyperlink } from "./links";
import { Paragraph } from "./paragraph";
import { parseParagraph } from "./paragraph-parse";
import { TextRun } from "./run";

describe("Paragraph", () => {
  beforeEach(() => {
    vi.spyOn(convenienceFunctions, "uniqueId").mockReturnValue("test-unique-id");
    vi.spyOn(convenienceFunctions, "bookmarkUniqueNumericIdGen").mockReturnValue(() => -101);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe("#constructor()", () => {
    it("should create valid JSON", () => {
      const paragraph = new Paragraph("");
      const stringifiedJson = JSON.stringify(paragraph);

      try {
        JSON.parse(stringifiedJson);
      } catch {
        assert.isTrue(false);
      }
      assert.isTrue(true);
    });

    it("should create have valid properties", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.LEFT,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.have.property("w:p").which.is.an("array");
      expect(tree["w:p"][0]).to.have.property("w:pPr");
    });
  });

  describe("#heading1()", () => {
    it("should add heading style to JSON", () => {
      const paragraph = new Paragraph({
        heading: HeadingLevel.HEADING_1,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "Heading1" } } }],
          },
        ],
      });
    });
  });

  describe("#heading2()", () => {
    it("should add heading style to JSON", () => {
      const paragraph = new Paragraph({
        heading: HeadingLevel.HEADING_2,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "Heading2" } } }],
          },
        ],
      });
    });
  });

  describe("#heading3()", () => {
    it("should add heading style to JSON", () => {
      const paragraph = new Paragraph({
        heading: HeadingLevel.HEADING_3,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "Heading3" } } }],
          },
        ],
      });
    });
  });

  describe("#heading4()", () => {
    it("should add heading style to JSON", () => {
      const paragraph = new Paragraph({
        heading: HeadingLevel.HEADING_4,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "Heading4" } } }],
          },
        ],
      });
    });
  });

  describe("#heading5()", () => {
    it("should add heading style to JSON", () => {
      const paragraph = new Paragraph({
        heading: HeadingLevel.HEADING_5,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "Heading5" } } }],
          },
        ],
      });
    });
  });

  describe("#heading6()", () => {
    it("should add heading style to JSON", () => {
      const paragraph = new Paragraph({
        heading: HeadingLevel.HEADING_6,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "Heading6" } } }],
          },
        ],
      });
    });
  });

  describe("#title()", () => {
    it("should add title style to JSON", () => {
      const paragraph = new Paragraph({
        heading: HeadingLevel.TITLE,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "Title" } } }],
          },
        ],
      });
    });
  });

  describe("#center()", () => {
    it("should add center alignment to JSON", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.CENTER,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:jc": { _attr: { "w:val": "center" } } }],
          },
        ],
      });
    });
  });

  describe("#left()", () => {
    it("should add left alignment to JSON", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.LEFT,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:jc": { _attr: { "w:val": "left" } } }],
          },
        ],
      });
    });
  });

  describe("#right()", () => {
    it("should add right alignment to JSON", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.RIGHT,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:jc": { _attr: { "w:val": "right" } } }],
          },
        ],
      });
    });
  });

  describe("#start()", () => {
    it("should add start alignment to JSON", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.START,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:jc": { _attr: { "w:val": "start" } } }],
          },
        ],
      });
    });
  });

  describe("#end()", () => {
    it("should add end alignment to JSON", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.END,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:jc": { _attr: { "w:val": "end" } } }],
          },
        ],
      });
    });
  });

  describe("#distribute()", () => {
    it("should add distribute alignment to JSON", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.DISTRIBUTE,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:jc": { _attr: { "w:val": "distribute" } } }],
          },
        ],
      });
    });
  });

  describe("#justified()", () => {
    it("should add justified alignment to JSON", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:jc": { _attr: { "w:val": "both" } } }],
          },
        ],
      });
    });
  });

  describe("#maxRightTabStop()", () => {
    it("should add right tab stop to JSON", () => {
      const paragraph = new Paragraph({
        tabStops: [
          {
            position: TabStopPosition.MAX,
            type: TabStopType.RIGHT,
          },
        ],
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:tabs": [
                  {
                    "w:tab": {
                      _attr: {
                        "w:pos": 9026,
                        "w:val": "right",
                      },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  describe("#leftTabStop()", () => {
    it("should add leftTabStop to JSON", () => {
      const paragraph = new Paragraph({
        tabStops: [
          {
            leader: LeaderType.HYPHEN,
            position: 100,
            type: TabStopType.LEFT,
          },
        ],
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:tabs": [
                  {
                    "w:tab": {
                      _attr: {
                        "w:leader": "hyphen",
                        "w:pos": 100,
                        "w:val": "left",
                      },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  describe("#rightTabStop()", () => {
    it("should add rightTabStop to JSON", () => {
      const paragraph = new Paragraph({
        tabStops: [
          {
            leader: LeaderType.DOT,
            position: 100,
            type: TabStopType.RIGHT,
          },
        ],
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:tabs": [
                  {
                    "w:tab": {
                      _attr: {
                        "w:leader": "dot",
                        "w:pos": 100,
                        "w:val": "right",
                      },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  describe("#centerTabStop()", () => {
    it("should add centerTabStop to JSON", () => {
      const paragraph = new Paragraph({
        tabStops: [
          {
            leader: LeaderType.MIDDLE_DOT,
            position: 100,
            type: TabStopType.CENTER,
          },
        ],
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:tabs": [
                  {
                    "w:tab": {
                      _attr: {
                        "w:leader": "middleDot",
                        "w:pos": 100,
                        "w:val": "center",
                      },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  describe("#contextualSpacing()", () => {
    it("should add contextualSpacing", () => {
      const paragraph = new Paragraph({
        contextualSpacing: true,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:contextualSpacing": {} }],
          },
        ],
      });
    });
    it("should remove contextualSpacing", () => {
      const paragraph = new Paragraph({
        contextualSpacing: false,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:contextualSpacing": { _attr: { "w:val": false } } }],
          },
        ],
      });
    });
  });

  describe("#thematicBreak()", () => {
    it("should add thematic break to JSON", () => {
      const paragraph = new Paragraph({
        thematicBreak: true,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:pBdr": [
                  {
                    "w:bottom": {
                      _attr: {
                        "w:color": "auto",
                        "w:space": 1,
                        "w:sz": 6,
                        "w:val": "single",
                      },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  describe("#paragraphBorders()", () => {
    it("should add a left and right border to a paragraph", () => {
      const paragraph = new Paragraph({
        border: {
          left: {
            color: "auto",
            size: 6,
            space: 1,
            style: BorderStyle.SINGLE,
          },
          right: {
            color: "auto",
            size: 6,
            space: 1,
            style: BorderStyle.SINGLE,
          },
        },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:pBdr": [
                  {
                    "w:left": {
                      _attr: {
                        "w:color": "auto",
                        "w:space": 1,
                        "w:sz": 6,
                        "w:val": "single",
                      },
                    },
                  },
                  {
                    "w:right": {
                      _attr: {
                        "w:color": "auto",
                        "w:space": 1,
                        "w:sz": 6,
                        "w:val": "single",
                      },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  describe("#pageBreak()", () => {
    it("should add page break to JSON", () => {
      const paragraph = new Paragraph({
        children: [new PageBreak()],
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:r": [{ "w:br": { _attr: { "w:type": "page" } } }],
          },
        ],
      });
    });
  });

  describe("#pageBreakBefore()", () => {
    it("should add page break before to JSON", () => {
      const paragraph = new Paragraph({
        pageBreakBefore: true,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:pageBreakBefore": EMPTY_OBJECT,
              },
            ],
          },
        ],
      });
    });
  });

  describe("#bullet()", () => {
    it("should default to 0 indent level if no bullet was specified", () => {
      const paragraph = new Paragraph({
        bullet: {
          level: 0,
        },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.have.property("w:p").which.is.an("array").which.has.length.at.least(1);
      expect(tree["w:p"][0])
        .to.have.property("w:pPr")
        .which.is.an("array")
        .which.has.length.at.least(1);
      expect(tree["w:p"][0]["w:pPr"][0]).to.deep.equal({
        "w:pStyle": { _attr: { "w:val": "ListParagraph" } },
      });
    });

    it("should add list paragraph style to JSON", () => {
      const paragraph = new Paragraph({
        bullet: {
          level: 0,
        },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.have.property("w:p").which.is.an("array").which.has.length.at.least(1);
      expect(tree["w:p"][0])
        .to.have.property("w:pPr")
        .which.is.an("array")
        .which.has.length.at.least(1);
      expect(tree["w:p"][0]["w:pPr"][0]).to.deep.equal({
        "w:pStyle": { _attr: { "w:val": "ListParagraph" } },
      });
    });

    it("it should add numbered properties", () => {
      const paragraph = new Paragraph({
        bullet: {
          level: 1,
        },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.have.property("w:p").which.is.an("array").which.has.length.at.least(1);
      expect(tree["w:p"][0])
        .to.have.property("w:pPr")
        .which.is.an("array")
        .which.has.length.at.least(2);
      expect(tree["w:p"][0]["w:pPr"][1]).to.deep.equal({
        "w:numPr": [
          { "w:ilvl": { _attr: { "w:val": 1 } } },
          { "w:numId": { _attr: { "w:val": 1 } } },
        ],
      });
    });
  });

  describe("#setNumbering", () => {
    const createNumberingContext = () => ({
      file: {
        Numbering: {
          createConcreteNumberingInstance: (_: string, __: number) => undefined,
        },
      } as unknown as File,
      fileData: {
        Numbering: {
          createConcreteNumberingInstance: (_: string, __: number) => undefined,
        },
      } as unknown as File,
      stack: [],
      viewWrapper: new DocumentWrapper({ background: {} }),
    });

    it("should add list paragraph style to JSON", () => {
      const paragraph = new Paragraph({
        numbering: {
          level: 0,
          reference: "test id",
        },
      });
      const tree = new Formatter().format(paragraph, createNumberingContext());
      expect(tree).to.have.property("w:p").which.is.an("array").which.has.length.at.least(1);
      expect(tree["w:p"][0])
        .to.have.property("w:pPr")
        .which.is.an("array")
        .which.has.length.at.least(1);
      expect(tree["w:p"][0]["w:pPr"][0]).to.deep.equal({
        "w:pStyle": { _attr: { "w:val": "ListParagraph" } },
      });
    });

    it("should add a style to the list paragraph when provided", () => {
      const paragraph = new Paragraph({
        numbering: {
          level: 0,
          reference: "test id",
        },
        style: "myFancyStyle",
      });
      const tree = new Formatter().format(paragraph, createNumberingContext());
      expect(tree).to.have.property("w:p").which.is.an("array").which.has.length.at.least(1);
      expect(tree["w:p"][0])
        .to.have.property("w:pPr")
        .which.is.an("array")
        .which.has.length.at.least(1);
      expect(tree["w:p"][0]["w:pPr"][0]).to.deep.equal({
        "w:pStyle": { _attr: { "w:val": "myFancyStyle" } },
      });
    });

    it("should not add ListParagraph style to a list when using custom numbering", () => {
      const paragraph = new Paragraph({
        numbering: {
          custom: true,
          level: 0,
          reference: "test id",
        },
      });
      const tree = new Formatter().format(paragraph, createNumberingContext());
      expect(tree).to.have.property("w:p").which.is.an("array").which.has.length.at.least(1);
      expect(tree["w:p"][0])
        .to.have.property("w:pPr")
        .which.is.an("array")
        .which.has.length.at.least(1);
      expect(tree["w:p"][0]["w:pPr"][0]).to.not.have.property("w:pStyle");
    });

    it("it should add numbered properties", () => {
      const paragraph = new Paragraph({
        numbering: {
          instance: 4,
          level: 0,
          reference: "test id",
        },
      });
      const tree = new Formatter().format(paragraph, createNumberingContext());
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              { "w:pStyle": { _attr: { "w:val": "ListParagraph" } } },
              {
                "w:numPr": [
                  { "w:ilvl": { _attr: { "w:val": 0 } } },
                  { "w:numId": { _attr: { "w:val": "{test id-4}" } } },
                ],
              },
            ],
          },
        ],
      });
    });

    it("should not add ListParagraph style when custom is true", () => {
      const paragraph = new Paragraph({
        numbering: {
          custom: true,
          level: 0,
          reference: "test id",
        },
      });
      const tree = new Formatter().format(paragraph, createNumberingContext());
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:numPr": [
                  { "w:ilvl": { _attr: { "w:val": 0 } } },
                  { "w:numId": { _attr: { "w:val": "{test id-0}" } } },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  it("it should add bookmark", () => {
    const paragraph = new Paragraph({
      children: [
        new Bookmark({
          children: [new TextRun("test")],
          id: "test-id",
        }),
      ],
    });
    const tree = new Formatter().format(paragraph);
    expect(tree).to.deep.equal({
      "w:p": [
        {
          "w:bookmarkStart": {
            _attr: {
              "w:id": -101,
              "w:name": "test-id",
            },
          },
        },
        {
          "w:r": [
            {
              "w:t": [
                {
                  _attr: {
                    "xml:space": "preserve",
                  },
                },
                "test",
              ],
            },
          ],
        },
        {
          "w:bookmarkEnd": {
            _attr: {
              "w:id": -101,
            },
          },
        },
      ],
    });
  });

  describe("#style", () => {
    it("should set the paragraph style to the given styleId", () => {
      const paragraph = new Paragraph({
        style: "myFancyStyle",
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:pStyle": { _attr: { "w:val": "myFancyStyle" } } }],
          },
        ],
      });
    });
  });

  describe("#indent", () => {
    it("should set the paragraph indent to the given values", () => {
      const paragraph = new Paragraph({
        indent: { left: 720 },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:ind": { _attr: { "w:left": 720 } } }],
          },
        ],
      });
    });
  });

  describe("#spacing", () => {
    it("should set the paragraph spacing to the given values", () => {
      const paragraph = new Paragraph({
        spacing: { before: 90, line: 50 },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:spacing": { _attr: { "w:before": 90, "w:line": 50 } } }],
          },
        ],
      });
    });
  });

  describe("#keepLines", () => {
    it("should set the paragraph keepLines sub-component", () => {
      const paragraph = new Paragraph({
        keepLines: true,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [{ "w:pPr": [{ "w:keepLines": EMPTY_OBJECT }] }],
      });
    });
  });

  describe("#keepNext", () => {
    it("should set the paragraph keepNext sub-component", () => {
      const paragraph = new Paragraph({
        keepNext: true,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [{ "w:pPr": [{ "w:keepNext": EMPTY_OBJECT }] }],
      });
    });
  });

  describe("#bidirectional", () => {
    it("set paragraph right to left layout", () => {
      const paragraph = new Paragraph({
        bidirectional: true,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [{ "w:pPr": [{ "w:bidi": EMPTY_OBJECT }] }],
      });
    });
  });

  describe("#suppressLineNumbers", () => {
    it("should disable line numbers", () => {
      const paragraph = new Paragraph({
        suppressLineNumbers: true,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [{ "w:pPr": [{ "w:suppressLineNumbers": EMPTY_OBJECT }] }],
      });
    });
  });

  describe("#outlineLevel", () => {
    it("should set paragraph outline level to the given value", () => {
      const paragraph = new Paragraph({
        outlineLevel: 0,
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [{ "w:outlineLvl": { _attr: { "w:val": 0 } } }],
          },
        ],
      });
    });
  });

  describe("#shading", () => {
    it("should set shading to the given value", () => {
      const paragraph = new Paragraph({
        shading: {
          color: "00FFFF",
          fill: "FF0000",
          type: ShadingType.REVERSE_DIAGONAL_STRIPE,
        },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:shd": {
                  _attr: {
                    "w:color": "00FFFF",
                    "w:fill": "FF0000",
                    "w:val": "reverseDiagStripe",
                  },
                },
              },
            ],
          },
        ],
      });
    });
  });

  describe("#frame", () => {
    it("should set frame attribute", () => {
      const paragraph = new Paragraph({
        frame: {
          alignment: {
            x: HorizontalPositionAlign.CENTER,
            y: VerticalPositionAlign.TOP,
          },
          anchor: {
            horizontal: FrameAnchorType.MARGIN,
            vertical: FrameAnchorType.MARGIN,
          },
          height: 1000,
          type: "alignment",
          width: 4000,
        },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:framePr": {
                  _attr: {
                    "w:h": 1000,
                    "w:hAnchor": "margin",
                    "w:vAnchor": "margin",
                    "w:w": 4000,
                    "w:xAlign": "center",
                    "w:yAlign": "top",
                  },
                },
              },
            ],
          },
        ],
      });
    });
  });

  describe("#prepForXml", () => {
    it("should set Internal Hyperlink", () => {
      const paragraph = new Paragraph({
        children: [
          new ExternalHyperlink({
            children: [new TextRun("test")],
            link: "http://www.google.com",
          }),
        ],
      });
      const viewWrapperMock = {
        Relationships: {
          addRelationship: () => {},
        },
      } as unknown as ViewWrapper;

      const tree = new Formatter().format(paragraph, {
        file: {} as unknown as File,
        stack: [],
        viewWrapper: viewWrapperMock,
      });
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:hyperlink": [
              {
                _attr: {
                  "r:id": "rIdtest-unique-id",
                  "w:history": 1,
                },
              },
              {
                "w:r": [
                  {
                    "w:t": [
                      {
                        _attr: {
                          "xml:space": "preserve",
                        },
                      },
                      "test",
                    ],
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });

  describe("#revision", () => {
    it("should create paragraph with revision properties", () => {
      const paragraph = new Paragraph({
        alignment: AlignmentType.RIGHT,
        heading: HeadingLevel.HEADING_1,
        revision: {
          alignment: AlignmentType.LEFT,
          author: "Firstname Lastename",
          date: "123",
          heading: HeadingLevel.HEADING_2,
          id: 1,
        },
      });
      const tree = new Formatter().format(paragraph);
      expect(tree).to.deep.equal({
        "w:p": [
          {
            "w:pPr": [
              {
                "w:pStyle": {
                  _attr: {
                    "w:val": "Heading1",
                  },
                },
              },
              {
                "w:jc": {
                  _attr: {
                    "w:val": "right",
                  },
                },
              },
              {
                "w:pPrChange": [
                  {
                    _attr: {
                      "w:author": "Firstname Lastename",
                      "w:date": "123",
                      "w:id": 1,
                    },
                  },
                  {
                    "w:pPr": [
                      {
                        "w:pStyle": {
                          _attr: {
                            "w:val": "Heading2",
                          },
                        },
                      },
                      {
                        "w:jc": {
                          _attr: {
                            "w:val": "left",
                          },
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      });
    });
  });
});

// ── Parse round-trip tests ──────────────────────────────────────────────────

describe("parse round-trip", () => {
  // Minimal mock ParseContext (most parse functions only use _ctx)
  const ctx = new ParseContext(
    {
      body: {} as never,
      partRefs: {
        headers: new Map(),
        footers: new Map(),
        charts: new Map(),
        diagramData: new Map(),
        media: new Map(),
        afChunks: new Map(),
        subDocs: new Map(),
      },
    } as never,
    new Map(),
    new Map(),
  );

  function parseFormatted(component: Paragraph) {
    const tree = new Formatter().format(component);
    // Formatter returns { "w:p": [...] }, toElement gives us the root element
    const root = toElement(tree);
    // The root element IS the w:p element
    return parseParagraph(root, ctx);
  }

  it("should parse heading", () => {
    const paragraph = new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: "Title" })],
    });
    const parsed = parseFormatted(paragraph);
    expect(typeof parsed === "object").toBe(true);
    if (typeof parsed !== "object") return;
    expect(parsed.heading).toBe("Heading1");
  });

  it("should parse alignment", () => {
    const paragraph = new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Centered" })],
    });
    const parsed = parseFormatted(paragraph);
    expect(typeof parsed === "object").toBe(true);
    if (typeof parsed !== "object") return;
    expect(parsed.alignment).toBe("center");
  });

  it("should parse spacing", () => {
    const paragraph = new Paragraph({
      spacing: { before: 200, after: 100, line: 360 },
      children: [new TextRun({ text: "Spaced" })],
    });
    const parsed = parseFormatted(paragraph);
    expect(typeof parsed === "object").toBe(true);
    if (typeof parsed !== "object") return;
    expect(parsed.spacing).toBeDefined();
    const spacing = parsed.spacing as Record<string, unknown>;
    expect(spacing.before).toBe(200);
    expect(spacing.after).toBe(100);
    expect(spacing.line).toBe(360);
  });

  it("should parse indent", () => {
    const paragraph = new Paragraph({
      indent: { left: 720 },
      children: [new TextRun({ text: "Indented" })],
    });
    const parsed = parseFormatted(paragraph);
    expect(typeof parsed === "object").toBe(true);
    if (typeof parsed !== "object") return;
    expect(parsed.indent).toBeDefined();
    const indent = parsed.indent as Record<string, unknown>;
    expect(indent.left).toBe(720);
  });

  it("should parse simple text as string", () => {
    const paragraph = new Paragraph({ text: "Hello World" });
    const parsed = parseFormatted(paragraph);
    // Simple text-only paragraphs are optimized to string
    if (typeof parsed === "string") {
      expect(parsed).toBe("Hello World");
    } else {
      expect(parsed.text).toBe("Hello World");
    }
  });

  // ── JSON API coerce round-trip tests ───────────────────────────────────────

  it("should round-trip math via JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "E=mc" }, { math: { children: [{ text: "2" }] } } as never],
    });
    const tree = new Formatter().format(paragraph);
    // Should contain m:oMath element
    const root = toElement(tree);
    const mathEl = root.elements?.find((e) => e.name === "m:oMath");
    expect(mathEl).toBeDefined();
  });

  it("should round-trip symbolRun via JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "Arrow: " }, { symbolRun: { char: "F021", symbolfont: "Wingdings" } }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    // Find w:r containing w:sym
    const symRun = root.elements?.find((e) => e.elements?.some((c) => c.name === "w:sym"));
    expect(symRun).toBeDefined();
  });

  it("should round-trip pageBreak via JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "Before" }, { pageBreak: true }, { text: "After" }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    const brRun = root.elements?.find((e) =>
      e.elements?.some((c) => c.name === "w:br" && c.attributes?.["w:type"] === "page"),
    );
    expect(brRun).toBeDefined();
  });

  it("should round-trip columnBreak via JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "Col1" }, { columnBreak: true }, { text: "Col2" }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    const brRun = root.elements?.find((e) =>
      e.elements?.some((c) => c.name === "w:br" && c.attributes?.["w:type"] === "column"),
    );
    expect(brRun).toBeDefined();
  });

  it("should round-trip commentRangeStart/End/Reference via JSON API", () => {
    const paragraph = new Paragraph({
      children: [
        { commentRangeStart: 0 },
        { text: "Commented" },
        { commentRangeEnd: 0 },
        { commentReference: 0 },
      ],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    expect(root.elements?.some((e) => e.name === "w:commentRangeStart")).toBe(true);
    expect(root.elements?.some((e) => e.name === "w:commentRangeEnd")).toBe(true);
    // CommentReference is nested inside w:r (run-level element per OOXML spec)
    const refRun = root.elements?.find(
      (e) => e.name === "w:r" && e.elements?.some((c) => c.name === "w:commentReference"),
    );
    expect(refRun).toBeDefined();
  });

  it("should round-trip footnoteReference via JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "With note" }, { footnoteReference: 1 }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    // footnoteReference is inside a w:r
    const refRun = root.elements?.find(
      (e) => e.name === "w:r" && e.elements?.some((c) => c.name === "w:footnoteReference"),
    );
    expect(refRun).toBeDefined();
  });

  it("should round-trip endnoteReference via JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "With endnote" }, { endnoteReference: 1 }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    // endnoteReference is inside a w:r
    const refRun = root.elements?.find(
      (e) => e.name === "w:r" && e.elements?.some((c) => c.name === "w:endnoteReference"),
    );
    expect(refRun).toBeDefined();
  });
});
