// Image click and hover hyperlinks
import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                bold: true,
                size: 16,
                text: "Image Hyperlink Demo",
              },
            ],
            spacing: { after: 400 },
          },
        },

        // 1. Image with click hyperlink
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "1. Image with Click Hyperlink",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: ["(Click the image below to open https://example.com)"],
            spacing: { after: 100 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  altText: {
                    description: "Click me!",
                    hyperlink: { click: "https://example.com" },
                    name: "link-image",
                    title: "Click to visit example.com",
                  },
                  data: readFileSync("./demo/images/cat.jpg"),
                  transformation: {
                    height: 150,
                    width: 150,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Image with both click and hover hyperlink
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "2. Image with Click + Hover Hyperlink",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: ["(Click opens https://example.com, hover opens https://example.com/hover)"],
            spacing: { after: 100 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  altText: {
                    description: "Click or hover me!",
                    hyperlink: {
                      click: "https://example.com",
                      hover: "https://example.com/hover",
                    },
                    name: "dual-link-image",
                    title: "Dual hyperlink image",
                  },
                  data: readFileSync("./demo/images/cat.jpg"),
                  transformation: {
                    height: 150,
                    width: 150,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
