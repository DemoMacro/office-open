// Text Frame (Text Box) example

import { writeFileSync } from "node:fs";

import {
  AlignmentType,
  BorderStyle,
  FrameAnchorType,
  HorizontalPositionAlign,
  VerticalPositionAlign,
  generateDocument,
} from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            border: {
              bottom: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
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
              top: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
            },
            children: [
              "Hello World",
              {
                text: "Foo Bar",
                bold: true,
              },
              {
                children: [{ tab: true }, "Github is the best"],
                bold: true,
              },
            ],
            frame: {
              anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
              },
              height: "1.8cm",
              position: {
                x: "1.8cm",
                y: "5.3cm",
              },
              type: "absolute",
              width: "7.1cm",
            },
          },
        },
        {
          paragraph: {
            border: {
              bottom: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
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
              top: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
            },
            children: [
              "Hello World",
              {
                text: "Foo Bar",
                bold: true,
              },
              {
                children: [{ tab: true }, "Github is the best"],
                bold: true,
              },
            ],
            frame: {
              alignment: {
                x: HorizontalPositionAlign.CENTER,
                y: VerticalPositionAlign.TOP,
              },
              anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
              },
              height: "1.8cm",
              type: "alignment",
              width: "7.1cm",
            },
          },
        },
        {
          paragraph: {
            alignment: AlignmentType.RIGHT,
            border: {
              bottom: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
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
              top: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
            },
            children: [
              "Hello World",
              {
                text: "Foo Bar",
                bold: true,
              },
              {
                children: [{ tab: true }, "Github is the best"],
                bold: true,
              },
            ],
            frame: {
              alignment: {
                x: HorizontalPositionAlign.CENTER,
                y: VerticalPositionAlign.BOTTOM,
              },
              anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
              },
              height: "1.8cm",
              type: "alignment",
              width: "7.1cm",
            },
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
