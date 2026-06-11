// This example shows 3 styles using XML styles

import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument, HeadingLevel } from "@office-open/docx";

const styles = readFileSync("./demo/assets/custom-styles.xml", "utf8");
const buffer = await generateDocument({
  externalStyles: styles,
  sections: [
    {
      children: [
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "Cool Heading Text",
          },
        },
        {
          paragraph: {
            style: "MyFancyStyle",
            text: 'This is a custom named style from the template "MyFancyStyle"',
          },
        },
        { paragraph: "Some normal text" },
        {
          paragraph: {
            style: "MyFancyStyle",
            text: "MyFancyStyle again",
          },
        },
      ],
    },
  ],
  title: "Title",
});
writeFileSync("My Document.docx", buffer);
