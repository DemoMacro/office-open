import { readFileSync, writeFileSync } from "node:fs";
// Example of how to add images to the document - You can use Buffers, UInt8Arrays or Base64 strings

import {
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  VerticalPositionAlign,
  VerticalPositionRelativeFrom,
  generateDocument,
} from "@office-open/docx";

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
                  altText: {
                    description: "This is an ultimate image",
                    name: "My Ultimate Image",
                    title: "This is an ultimate title",
                  },
                  data: readFileSync("./demo/images/image1.jpeg"),
                  transformation: {
                    height: 100,
                    width: 100,
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
                  data: readFileSync("./demo/images/dog.png").toString("base64"),
                  outline: {
                    color: { value: "FF0000" },
                    type: "solidFill",
                  },
                  transformation: {
                    height: 100,
                    width: 100,
                  },
                  type: "png",
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
                  data: readFileSync("./demo/images/cat.jpg"),
                  outline: {
                    color: { value: "0000FF" },
                    type: "solidFill",
                    width: "600mm",
                  },
                  transformation: {
                    flip: {
                      vertical: true,
                    },
                    height: 100,
                    width: 100,
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
                  data: readFileSync("./demo/images/parrots.bmp"),
                  transformation: {
                    flip: {
                      horizontal: true,
                    },
                    height: 150,
                    rotation: 225,
                    width: 150,
                  },
                  type: "bmp",
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
                    flip: {
                      horizontal: true,
                      vertical: true,
                    },
                    height: 200,
                    width: 200,
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
                  floating: {
                    horizontalPosition: {
                      offset: 1_014_400,
                    },
                    verticalPosition: {
                      offset: 1_014_400,
                    },
                    zIndex: 10,
                  },
                  transformation: {
                    height: 200,
                    rotation: 45,
                    width: 200,
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
                  data: readFileSync("./demo/images/cat.jpg"),
                  floating: {
                    horizontalPosition: {
                      align: HorizontalPositionAlign.RIGHT,
                      relative: HorizontalPositionRelativeFrom.PAGE,
                    },
                    verticalPosition: {
                      align: VerticalPositionAlign.BOTTOM,
                      relative: VerticalPositionRelativeFrom.PAGE,
                    },
                    zIndex: 5,
                  },
                  transformation: {
                    height: 200,
                    width: 200,
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
                  data: readFileSync("./demo/images/linux-svg.svg"),
                  fallback: {
                    data: readFileSync("./demo/images/linux-png.png"),
                    type: "png",
                  },
                  transformation: {
                    height: 200,
                    width: 200,
                  },
                  type: "svg",
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
