// This example shows 3 styles using XML styles

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { generateDocument, HeadingLevel } from "@office-open/docx";

const styles = readFileSync("./demo/assets/custom-styles.xml", "utf8");
const buffer = await generateDocument({
  styles: { external: styles },
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
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/13-xml-styles.docx", buffer);
