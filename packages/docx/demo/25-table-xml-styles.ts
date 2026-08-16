// Example of how you would create a table and add data to it

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { WidthType, generateDocument } from "@office-open/docx";

const styles = readFileSync("./demo/assets/custom-styles.xml", "utf8");

const buffer = await generateDocument({
  styles: { external: styles },
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
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/25-table-xml-styles.docx", buffer);
