// Simple example to add text to a document

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    "2+2",
                    {
                      fraction: {
                        denominator: ["2"],
                        numerator: ["hi"],
                      },
                    },
                  ],
                },
              },
              {
                bold: true,
                text: "Foo Bar",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      fraction: {
                        denominator: ["2"],
                        numerator: ["1", { radical: { children: ["2"] } }],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    { sum: { children: ["test"] } },
                    {
                      sum: {
                        children: [{ superScript: { children: ["e"], superScript: ["2"] } }],
                        subScript: ["i"],
                      },
                    },
                    {
                      sum: {
                        children: [{ radical: { children: ["i"] } }],
                        subScript: ["i"],
                        superScript: ["10"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    { integral: { children: ["test"] } },
                    {
                      integral: {
                        children: [{ superScript: { children: ["e"], superScript: ["2"] } }],
                        subScript: ["i"],
                      },
                    },
                    {
                      integral: {
                        children: [{ radical: { children: ["i"] } }],
                        subScript: ["i"],
                        superScript: ["10"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [{ superScript: { children: ["test"], superScript: ["hello"] } }],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [{ subScript: { children: ["test"], subScript: ["hello"] } }],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      subScript: {
                        children: ["x"],
                        subScript: [{ superScript: { children: ["y"], superScript: ["2"] } }],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      subSuperScript: {
                        children: ["test"],
                        subScript: ["world"],
                        superScript: ["hello"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      preSubSuperScript: {
                        children: ["test"],
                        subScript: ["world"],
                        superScript: ["hello"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      subScript: {
                        children: [{ fraction: { denominator: ["2"], numerator: ["1"] } }],
                        subScript: ["4"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      subScript: {
                        children: [
                          {
                            radical: {
                              children: [{ fraction: { denominator: ["2"], numerator: ["1"] } }],
                              degree: ["4"],
                            },
                          },
                        ],
                        subScript: ["x"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [{ radical: { children: ["4"] } }],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      function: {
                        children: ["100"],
                        name: [{ superScript: { children: ["cos"], superScript: ["-1"] } }],
                      },
                    },
                    "×",
                    {
                      function: {
                        children: ["360"],
                        name: ["sin"],
                      },
                    },
                    "= x",
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      roundBrackets: [{ fraction: { denominator: ["2"], numerator: ["1"] } }],
                    },
                    {
                      squareBrackets: [{ fraction: { denominator: ["2"], numerator: ["1"] } }],
                    },
                    {
                      curlyBrackets: [{ fraction: { denominator: ["2"], numerator: ["1"] } }],
                    },
                    {
                      angledBrackets: [{ fraction: { denominator: ["2"], numerator: ["1"] } }],
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      fraction: {
                        denominator: ["2a"],
                        numerator: [{ radical: { children: ["4"] } }],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    { limitUpper: { children: ["x"], limit: ["-"] } },
                    "=",
                    { limitLower: { children: ["lim"], limit: ["x→0"] } },
                  ],
                },
              },
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
