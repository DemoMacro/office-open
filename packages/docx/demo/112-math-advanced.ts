// Demo: Advanced math elements - box, borderBox, eqArr, groupChr, matrix, phantom
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Advanced Math Elements Demo", bold: true, size: 32 }],
            spacing: { after: 400 },
          },
        },

        // 1. MathBox
        {
          paragraph: {
            children: [{ bold: true, text: "1. MathBox", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    { box: { children: ["x + y"], properties: { opEmu: true } } },
                    " = ",
                    { box: { children: ["z"] } },
                  ],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. MathBorderBox
        {
          paragraph: {
            children: [{ bold: true, text: "2. MathBorderBox", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    { borderBox: { children: ["a"] } },
                    " + ",
                    {
                      borderBox: {
                        children: ["b"],
                        properties: { hideTop: true, hideBottom: true },
                      },
                    },
                  ],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Equation Array
        {
          paragraph: {
            children: [{ bold: true, text: "3. Equation Array", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [{ eqArr: { rows: [["x + y = 1"], ["2x - y = 3"]] } }],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 4. Group Character
        {
          paragraph: {
            children: [{ bold: true, text: "4. Group Character (brace)", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      groupChr: {
                        children: [{ eqArr: { rows: [["a"], ["b"], ["c"]] } }],
                        properties: {
                          chr: "{",
                          pos: "bot",
                          vertJc: "top",
                        },
                      },
                    },
                  ],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 5. Matrix (with column properties: baseJc, plcHide, rSpRule)
        {
          paragraph: {
            children: [{ bold: true, text: "5. Matrix", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      matrix: {
                        rows: [
                          ["1", "0"],
                          ["0", "1"],
                        ],
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
              "With properties: ",
              {
                math: {
                  children: [
                    {
                      matrix: {
                        rows: [
                          ["a", "b"],
                          ["c", "d"],
                        ],
                        properties: {
                          baseJc: "center",
                          plcHide: true,
                          rSpRule: 2,
                        },
                      },
                    },
                  ],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 6. Phantom
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "6. Phantom (invisible placeholder)",
                size: 28,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    "f(x) = ",
                    {
                      fraction: {
                        numerator: [
                          {
                            phant: {
                              children: ["dy"],
                              properties: { zeroAsc: true, zeroDesc: true },
                            },
                          },
                        ],
                        denominator: ["dx"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 7. SuperScript
        {
          paragraph: {
            children: [{ bold: true, text: "7. SuperScript (E = mc²)", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: ["E = m", { superScript: { children: ["c"], superScript: ["2"] } }],
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
