// Demo: Ruby annotation (CT_Ruby) - East Asian pronunciation guides
import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const rubyProperties = {
  alignment: "center" as const,
  fontSize: 10,
  raise: 10,
  baseFontSize: 20,
  languageId: "ja-JP",
};

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Ruby Annotation Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },
        {
          paragraph: {
            children: [
              "Japanese furigana: ",
              {
                children: [
                  {
                    ruby: {
                      properties: rubyProperties,
                      text: { children: [{ text: "かな", size: 10 }] },
                      base: { children: [{ text: "漢字", size: 20 }] },
                    },
                  },
                ],
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Chinese pinyin: ",
              {
                children: [
                  {
                    ruby: {
                      properties: { ...rubyProperties, languageId: "zh-CN" },
                      text: { children: ["hàn zì"] },
                      base: { children: ["汉字"] },
                    },
                  },
                ],
              },
            ],
          },
        },
      ],
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/107-ruby-annotation.docx", buffer);
