// Demo: Ruby annotation (CT_Ruby) - East Asian pronunciation guides
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Ruby Annotation Demo", bold: true, size: 32 }],
            spacing: { after: 400 },
          },
        },

        // Japanese furigana
        {
          paragraph: {
            children: [{ bold: true, text: "Japanese Furigana", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "Base text with ruby: ",
              {
                ruby: {
                  text: "かな",
                  base: "漢字",
                  alignment: "center",
                  fontSize: 20,
                  raise: 20,
                  baseFontSize: 40,
                  languageId: "ja-JP",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // Chinese pinyin
        {
          paragraph: {
            children: [{ bold: true, text: "Chinese Pinyin", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "Pinyin guide: ",
              {
                ruby: {
                  text: "hàn zì",
                  base: "汉字",
                  alignment: "center",
                  fontSize: 18,
                  raise: 20,
                  baseFontSize: 40,
                  languageId: "zh-CN",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // Alignment options
        {
          paragraph: {
            children: [{ bold: true, text: "Alignment Options", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "Left: ",
              { ruby: { text: "left", base: "Align", alignment: "left" } },
              "  Center: ",
              { ruby: { text: "center", base: "Align", alignment: "center" } },
              "  Right: ",
              { ruby: { text: "right", base: "Align", alignment: "right" } },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
