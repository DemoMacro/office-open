// Demo: Structured Document Tags (SDT) - content controls
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "SDT Content Controls Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },

        // Plain text SDT (block-level)
        {
          paragraph: {
            children: [{ bold: true, text: "1. Plain Text SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Name",
              tag: "name",
              text: { multiLine: false },
            },
            children: [{ paragraph: { children: ["John Doe"] } }],
          },
        },

        { paragraph: { children: [""] } },

        // Multi-line text SDT (block-level)
        {
          paragraph: {
            children: [{ bold: true, text: "2. Multi-line Text SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Description",
              tag: "multiline-text",
              text: { multiLine: true },
            },
            children: [{ paragraph: { children: ["This is a multi-line text content control."] } }],
          },
        },

        { paragraph: { children: [""] } },

        // ComboBox SDT (block-level)
        {
          paragraph: {
            children: [{ bold: true, text: "3. ComboBox SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Color",
              tag: "color",
              comboBox: {
                items: [
                  { displayText: "Red", value: "red" },
                  { displayText: "Blue", value: "blue" },
                  { displayText: "Green", value: "green" },
                ],
                lastValue: "Red",
              },
            },
            children: [{ paragraph: { children: ["Red"] } }],
          },
        },

        { paragraph: { children: [""] } },

        // DropDownList SDT (block-level)
        {
          paragraph: {
            children: [{ bold: true, text: "4. DropDownList SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Priority",
              tag: "priority",
              dropDownList: {
                items: [{ displayText: "High" }, { displayText: "Medium" }, { displayText: "Low" }],
                lastValue: "Medium",
              },
            },
            children: [{ paragraph: { children: ["Medium"] } }],
          },
        },

        { paragraph: { children: [""] } },

        // Date SDT (block-level)
        {
          paragraph: {
            children: [{ bold: true, text: "5. Date SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "DueDate",
              tag: "due-date",
              date: {
                dateFormat: "yyyy-MM-dd",
                languageId: "en-US",
                fullDate: "2026-04-29T00:00:00",
              },
            },
            children: [{ paragraph: { children: ["2026-04-29"] } }],
          },
        },

        { paragraph: { children: [""] } },

        // Block-level SDT (section child)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "6. Block-level SDT (section child)",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              richText: true,
              alias: "BlockContent",
              tag: "block-content",
            },
            children: [
              { paragraph: { children: ["This is a block-level content control."] } },
              {
                paragraph: {
                  children: ["It can contain multiple paragraphs and tables."],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // SDT wrapping a table (block-level)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "7. SDT wrapping a table (block-level)",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Controlled Table",
              tag: "table-sdt",
              richText: true,
            },
            children: [
              {
                table: {
                  columnWidths: [3000, 3000],
                  rows: [
                    {
                      cells: [
                        { children: [{ paragraph: "Normal cell" }] },
                        { children: [{ paragraph: "Controlled cell" }] },
                      ],
                    },
                  ],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // SDT wrapping multiple paragraphs (group-like)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "8. SDT wrapping multiple blocks",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Controlled Section",
              tag: "section-sdt",
              richText: true,
            },
            children: [
              { paragraph: { children: ["First paragraph in controlled section."] } },
              {
                table: {
                  columnWidths: [3000, 3000],
                  rows: [
                    {
                      cells: [
                        { children: [{ paragraph: "Controlled row - cell 1" }] },
                        { children: [{ paragraph: "Controlled row - cell 2" }] },
                      ],
                    },
                  ],
                },
              },
              { paragraph: { children: ["Last paragraph in controlled section."] } },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // Checkbox SDT (Word 2010+ content control, w14:checkbox)
        // Click to toggle directly — no document protection required.
        {
          paragraph: {
            children: [{ bold: true, text: "9. Checkbox SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Subscribe",
              tag: "subscribe-checkbox",
              checkbox: { checked: false },
            },
          },
        },
        {
          sdt: {
            properties: { alias: "Accept Terms", tag: "accept-terms", checkbox: { checked: true } },
          },
        },

        { paragraph: { children: [""] } },

        // 10. Inline (run-level) SDT — CT_SdtRun
        //     Content controls embedded inline in a paragraph run, mixed with text.
        {
          paragraph: {
            children: [{ bold: true, text: "10. Inline (run-level) SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "Signed by: ",
              {
                sdt: {
                  properties: { alias: "InlineName", tag: "inline-name", text: {} },
                  children: ["Jane Doe"],
                },
              },
              "   Status: ",
              {
                sdt: {
                  properties: { alias: "InlineStatus", tag: "inline-status", richText: true },
                  children: [{ text: "Approved", bold: true, color: "008000" }],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 11. Cell-level SDT — CT_SdtCell (wraps a single table cell)
        {
          paragraph: {
            children: [{ bold: true, text: "11. Cell-level SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          table: {
            columnWidths: [3000, 3000],
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Normal cell" }] },
                  {
                    sdt: {
                      properties: { alias: "CellControl", tag: "cell-control", text: {} },
                      cells: [{ children: [{ paragraph: "Controlled cell" }] }],
                    },
                  },
                ],
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 12. Row-level SDT — CT_SdtRow (wraps an entire table row)
        {
          paragraph: {
            children: [{ bold: true, text: "12. Row-level SDT", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          table: {
            columnWidths: [3000, 3000],
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Header A" }] },
                  { children: [{ paragraph: "Header B" }] },
                ],
              },
              {
                sdt: {
                  properties: { alias: "RowControl", tag: "row-control", richText: true },
                  rows: [
                    {
                      cells: [
                        { children: [{ paragraph: "Controlled row - A" }] },
                        { children: [{ paragraph: "Controlled row - B" }] },
                      ],
                    },
                  ],
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
