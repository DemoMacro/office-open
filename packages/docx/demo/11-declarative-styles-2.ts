import { readFileSync, writeFileSync } from "node:fs";
// Setting styles with JavaScript configuration

import {
  AlignmentType,
  HeadingLevel,
  LevelFormat,
  TabStopPosition,
  UnderlineType,
  generateDocument,
} from "@office-open/docx";

const table = {
  table: {
    rows: [
      {
        cells: [
          {
            children: [{ paragraph: "Test cell 1." }],
          },
        ],
      },
      {
        cells: [
          {
            children: [{ paragraph: "Test cell 2." }],
          },
        ],
      },
      {
        cells: [
          {
            children: [{ paragraph: "Test cell 3." }],
          },
        ],
      },
      {
        cells: [
          {
            children: [{ paragraph: "Test cell 4." }],
          },
        ],
      },
    ],
  },
};

const buffer = await generateDocument({
  numbering: {
    config: [
      {
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1)",
            start: 50,
          },
        ],
        reference: "ref1",
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    width: 100,
                    height: 100,
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            text: "HEADING",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
          },
        },
        {
          paragraph: {
            text: "Ref. :",
            style: "normalPara",
          },
        },
        {
          paragraph: {
            text: "Date :",
            style: "normalPara",
          },
        },
        {
          paragraph: {
            text: "To,",
            style: "normalPara",
          },
        },
        {
          paragraph: {
            text: "The Superindenting Engineer,(O &M)",
            style: "normalPara",
          },
        },
        {
          paragraph: {
            text: "Sub : ",
            style: "normalPara",
          },
        },
        {
          paragraph: {
            text: "Ref. : ",
            style: "normalPara",
          },
        },
        {
          paragraph: {
            text: "Sir,",
            style: "normalPara",
          },
        },
        {
          paragraph: {
            text: "BRIEF DESCRIPTION",
            style: "normalPara",
          },
        },
        table,
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    width: 100,
                    height: 100,
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            text: "Test",
            style: "normalPara2",
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    width: 100,
                    height: 100,
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            text: "Test 2",
            style: "normalPara2",
          },
        },
        {
          paragraph: {
            text: "Numbered paragraph that has numbering attached to custom styles",
            style: "numberedPara",
          },
        },
        {
          paragraph: {
            text: "Numbered para would show up in the styles pane at Word",
            style: "numberedPara",
          },
        },
      ],
      footers: {
        default: [
          {
            paragraph: {
              text: "1",
              style: "normalPara",
              alignment: AlignmentType.RIGHT,
            },
          },
        ],
      },
      properties: {
        page: {
          margin: {
            bottom: 700,
            left: 700,
            right: 700,
            top: 700,
          },
        },
      },
    },
  ],
  styles: {
    default: {
      heading1: {
        paragraph: {
          alignment: AlignmentType.CENTER,
          spacing: { line: 340 },
        },
        run: {
          bold: true,
          color: "000000",
          font: "Calibri",
          size: 26,
          underline: {
            color: "000000",
            type: UnderlineType.SINGLE,
          },
        },
      },
      heading2: {
        paragraph: {
          spacing: { line: 340 },
        },
        run: {
          bold: true,
          font: "Calibri",
          size: 13,
        },
      },
      heading3: {
        paragraph: {
          spacing: { line: 276 },
        },
        run: {
          bold: true,
          font: "Calibri",
          size: 13,
        },
      },
      heading4: {
        paragraph: {
          alignment: AlignmentType.JUSTIFIED,
        },
        run: {
          bold: true,
          font: "Calibri",
          size: 13,
        },
      },
    },
    paragraphStyles: [
      {
        basedOn: "Normal",
        id: "normalPara",
        name: "Normal Para",
        next: "Normal",
        paragraph: {
          leftTabStop: 454,
          rightTabStop: TabStopPosition.MAX,
          spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
        },
        quickFormat: true,
        run: {
          bold: true,
          font: "Calibri",
          size: 13,
        },
      },
      {
        basedOn: "Normal",
        id: "normalPara2",
        name: "Normal Para2",
        next: "Normal",
        paragraph: {
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
        },
        quickFormat: true,
        run: {
          font: "Calibri",
          size: 13,
        },
      },
      {
        basedOn: "Normal",
        id: "aside",
        name: "Aside",
        next: "Normal",
        paragraph: {
          indent: { left: "0.5in" },
          spacing: { line: 276 },
        },
        run: {
          color: "999999",
          italic: true,
        },
      },
      {
        basedOn: "Normal",
        id: "wellSpaced",
        name: "Well Spaced",
        paragraph: {
          spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
        },
      },
      {
        basedOn: "Normal",
        id: "numberedPara",
        name: "Numbered Para",
        next: "Normal",
        paragraph: {
          leftTabStop: 454,
          numbering: {
            instance: 0,
            level: 0,
            reference: "ref1",
          },
          rightTabStop: TabStopPosition.MAX,
          spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
        },
        quickFormat: true,
        run: {
          bold: true,
          font: "Calibri",
          size: 13,
        },
      },
    ],
  },
});
writeFileSync("My Document.docx", buffer);
