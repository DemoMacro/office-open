// Example on how to add hyperlinks to websites

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  footnotes: {
    1: {
      children: [
        {
          children: [
            "Click here for the ",
            {
              hyperlink: {
                link: "http://www.example.com",
                children: ["Footnotes external hyperlink"],
              },
            },
          ],
        },
      ],
    },
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              { hyperlink: { link: "http://www.example.com", children: ["Anchor Text"] } },
              { footnoteReference: 1 },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                hyperlink: {
                  link: "http://www.google.com",
                  children: ["Google Image Link"],
                },
              },
              { hyperlink: { link: "https://www.bbc.co.uk/news", children: ["BBC News Link"] } },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "This is a hyperlink with formatting: ",
              },
              {
                hyperlink: {
                  link: "http://www.example.com",
                  children: [
                    { text: "A ", style: "Hyperlink" },
                    { text: "single ", bold: true, style: "Hyperlink" },
                    { text: "link", doubleStrike: true, style: "Hyperlink" },
                    { text: "1", superScript: true, style: "Hyperlink" },
                    { text: "!", style: "Hyperlink" },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Tooltip example: ",
              {
                hyperlink: {
                  link: "http://www.example.com",
                  children: ["Hover for tooltip"],
                  tooltip: "This is a tooltip shown on hover",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Target frame example: ",
              {
                hyperlink: {
                  link: "http://www.example.com",
                  children: ["Open in new window"],
                },
              },
            ],
          },
        },
      ],
      footers: {
        default: [
          {
            paragraph: {
              children: [
                "Click here for the ",
                {
                  hyperlink: {
                    link: "http://www.example.com",
                    children: ["Footer external hyperlink"],
                  },
                },
              ],
            },
          },
        ],
      },
      headers: {
        default: [
          {
            paragraph: {
              children: [
                "Click here for the ",
                {
                  hyperlink: {
                    link: "http://www.google.com",
                    children: ["Header external hyperlink"],
                  },
                },
              ],
            },
          },
        ],
      },
    },
  ],
  styles: {
    default: {
      hyperlink: {
        run: {
          color: "FF0000",
          underline: {
            color: "0000FF",
          },
        },
      },
    },
  },
});
writeFileSync("My Document.docx", buffer);
