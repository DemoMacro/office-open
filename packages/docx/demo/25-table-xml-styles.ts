// Example of how you would create a table and add data to it

import { readFileSync, writeFileSync } from "node:fs";

import { WidthType, generateDocument } from "@office-open/docx";

const styles = readFileSync("./demo/assets/custom-styles.xml", "utf8");

const buffer = await generateDocument({
  externalStyles: styles,
  sections: [
    {
      children: [
        {
          table: {
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Header Colum 1" }] },
                  { children: [{ paragraph: "Header Colum 2" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Column Content 3" }] },
                  { children: [{ paragraph: "Column Content 2" }] },
                ],
              },
            ],
            style: "MyCustomTableStyle",
            width: {
              size: 9070,
              type: WidthType.DXA,
            },
          },
        },
      ],
    },
  ],
  title: "Title",
});
writeFileSync("My Document.docx", buffer);
