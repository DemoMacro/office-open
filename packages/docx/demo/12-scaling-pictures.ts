// Scaling images

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        { paragraph: "Hello World" },
        {
          paragraph: {
            children: [
              {
                picture: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: "1.3cm",
                    width: "1.3cm",
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                picture: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: "2.6cm",
                    width: "2.6cm",
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                picture: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: "6.6cm",
                    width: "6.6cm",
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                picture: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: "10.6cm",
                    width: "10.6cm",
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
      ],
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/12-scaling-pictures.docx", buffer);
