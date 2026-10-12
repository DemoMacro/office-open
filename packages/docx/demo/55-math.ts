// Simple example to add text to a document

import { mkdirSync, writeFileSync } from "node:fs";

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
                  equations: [
                    [
                      "2+2",
                      {
                        fraction: {
                          denominator: ["2"],
                          numerator: ["hi"],
                        },
                      },
                    ],
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
                  equations: [
                    [
                      {
                        fraction: {
                          denominator: ["2"],
                          numerator: ["1", { radical: { children: ["2"] } }],
                        },
                      },
                    ],
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
                  equations: [
                    [
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
                  equations: [
                    [
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
                  equations: [[{ superScript: { children: ["test"], superScript: ["hello"] } }]],
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
                  equations: [[{ subScript: { children: ["test"], subScript: ["hello"] } }]],
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
                  equations: [
                    [
                      {
                        subScript: {
                          children: ["x"],
                          subScript: [{ superScript: { children: ["y"], superScript: ["2"] } }],
                        },
                      },
                    ],
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
                  equations: [
                    [
                      {
                        subSuperScript: {
                          children: ["test"],
                          subScript: ["world"],
                          superScript: ["hello"],
                        },
                      },
                    ],
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
                  equations: [
                    [
                      {
                        preSubSuperScript: {
                          children: ["test"],
                          subScript: ["world"],
                          superScript: ["hello"],
                        },
                      },
                    ],
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
                  equations: [
                    [
                      {
                        subScript: {
                          children: [{ fraction: { denominator: ["2"], numerator: ["1"] } }],
                          subScript: ["4"],
                        },
                      },
                    ],
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
                  equations: [
                    [
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
                  equations: [[{ radical: { children: ["4"] } }]],
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
                  equations: [
                    [
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
                  equations: [
                    [
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
                  equations: [
                    [
                      {
                        fraction: {
                          denominator: ["2a"],
                          numerator: [{ radical: { children: ["4"] } }],
                        },
                      },
                    ],
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
                  equations: [
                    [
                      { limitUpper: { children: ["x"], limit: ["-"] } },
                      "=",
                      { limitLower: { children: ["lim"], limit: ["x→0"] } },
                    ],
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
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/55-math.docx", buffer);
