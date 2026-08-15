import * as fs from "fs";

import type {
  ColorDefinitionOptions,
  LayoutDefinitionOptions,
  StyleDefinitionOptions,
} from "@office-open/core/smartart";
import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

/**
 * A custom vertical process layout: rounded-rectangle steps stacked in a
 * snake column with a title node. Style labels reference this layout's
 * @styleLbl slots ("step", "title") from the custom quick style and colors.
 */
const customLayout: LayoutDefinitionOptions = {
  uniqueId: "urn:microsoft.com/office/officeart/2005/8/layout/customSteps",
  titles: [{ lang: "en-US", val: "Custom Steps" }],
  descriptions: [{ lang: "en-US", val: "A custom vertical process layout" }],
  categories: [{ type: "process", pri: 500 }],
  layoutNode: {
    name: "diagram",
    children: [
      {
        algorithm: {
          type: "snake",
          parameters: [
            { type: "grDir", value: "tL" },
            { type: "flowDir", value: "row" },
            { type: "contDir", value: "sameDir" },
            { type: "off", value: "ctr" },
          ],
        },
      },
      {
        constraints: [
          { type: "h", for: "ch", referenceType: "w", factor: 0.5 },
          { type: "sp", referenceType: "w", factor: 0.1 },
          { type: "primFontSz", for: "ch", operation: "equ", value: 60 },
        ],
      },
      {
        forEach: {
          axis: "ch",
          pointType: "node",
          children: [
            {
              layoutNode: {
                name: "node",
                styleLabel: "step",
                children: [
                  { algorithm: { type: "tx" } },
                  { shape: { type: "roundRect", adjustments: [{ idx: 1, val: 0.25 }] } },
                  { presentationOf: { axis: "desOrSelf", pointType: "node" } },
                  {
                    constraints: [
                      { type: "lMarg", referenceType: "primFontSz", factor: 0.3 },
                      { type: "rMarg", referenceType: "primFontSz", factor: 0.3 },
                      { type: "tMarg", referenceType: "primFontSz", factor: 0.25 },
                      { type: "bMarg", referenceType: "primFontSz", factor: 0.25 },
                    ],
                  },
                ],
              },
            },
          ],
        },
      },
    ],
  },
};

/** Quick style filling the custom layout's "step" slot with accent-themed fills. */
const customStyle: StyleDefinitionOptions = {
  uniqueId: "urn:microsoft.com/office/officeart/2005/8/quickstyle/customSteps",
  titles: [{ lang: "en-US", val: "Custom Steps Style" }],
  categories: [{ type: "simple", pri: 10200 }],
  styleLabels: [
    {
      name: "step",
      style: {
        lineReference: { index: 2, color: { value: "tx1" } },
        fillReference: { index: 1, color: { value: "accent1" } },
        effectReference: { index: 0, color: { value: "tx1" } },
        fontReference: { collection: "minor", color: { value: "tx1" } },
      },
    },
  ],
};

/** Color transform cycling accents through the "step" slot. */
const customColors: ColorDefinitionOptions = {
  uniqueId: "urn:microsoft.com/office/officeart/2005/8/colors/customSteps",
  titles: [{ lang: "en-US", val: "Custom Steps Colors" }],
  styleLabels: [
    {
      name: "step",
      fillColorList: {
        meth: "cycle",
        colors: [{ value: "accent1" }, { value: "accent3" }, { value: "accent5" }],
      },
      lineColorList: { colors: [{ value: "tx1", transforms: { tint: 75 } }] },
      textFillColorList: { colors: [{ value: "tx1" }] },
    },
  ],
};

const options: PresentationOptions = {
  title: "Custom SmartArt Layout Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "21.2cm",
            height: "1.3cm",
            textBody: {
              paragraphs: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Custom Layout / Style / Colors", size: 32, bold: true }],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "6.4cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [{ text: "Plan" }, { text: "Build" }, { text: "Ship" }, { text: "Iterate" }],
            layout: customLayout,
            style: customStyle,
            color: customColors,
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/35-smartart-custom-layout.pptx", buffer);
