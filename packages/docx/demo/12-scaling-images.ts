// Scaling images

import { readFileSync, writeFileSync } from "node:fs";

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
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: 50,
                    width: 50,
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
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: 100,
                    width: 100,
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
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: 250,
                    width: 250,
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
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: 400,
                    width: 400,
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
writeFileSync("My Document.docx", buffer);
