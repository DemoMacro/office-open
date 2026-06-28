// Add images to header and footer

import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [{ paragraph: "Hello World" }],
      headers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  image: {
                    data: readFileSync("./demo/images/image1.jpeg"),
                    transformation: {
                      height: "2.6cm",
                      width: "2.6cm",
                    },
                    type: "jpg",
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
                  image: {
                    data: readFileSync("./demo/images/image1.jpeg"),
                    transformation: {
                      height: "2.6cm",
                      width: "2.6cm",
                    },
                    type: "jpg",
                  },
                },
              ],
            },
          },
        ],
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
