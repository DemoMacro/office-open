// Demo: Checkboxes — static symbols, form-field checkboxes, and SDT content-control checkboxes
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      properties: {},
      children: [
        {
          paragraph: {
            children: [{ text: "Checkbox Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },

        // 1. Static checkbox symbols (via symbolRun — NOT interactive)
        {
          paragraph: {
            children: [{ bold: true, text: "1. Static symbols", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              { symbolRun: { char: "2610", symbolfont: "MS Gothic" } }, // ☐ unchecked
              "   ",
              { symbolRun: { char: "2612", symbolfont: "MS Gothic" } }, // ☒ checked
              { text: "   (not interactive)", italic: true },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Form-field checkboxes (w:checkBox — legacy form fields)
        //    Toggle by enabling "Restrict Editing" in Word.
        {
          paragraph: {
            children: [{ bold: true, text: "2. Form-field checkboxes", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              { formField: { name: "Check1", checkBox: { checked: false, sizeAuto: true } } },
              " Unchecked",
            ],
          },
        },
        {
          paragraph: {
            children: [
              { formField: { name: "Check2", checkBox: { checked: true, sizeAuto: true } } },
              " Checked",
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. SDT content-control checkboxes (w14:checkbox — Word 2010+)
        //    Click to toggle directly, no document protection required.
        {
          paragraph: {
            children: [{ bold: true, text: "3. Content-control checkboxes", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: { alias: "Subscribe", tag: "subscribe", checkbox: { checked: false } },
          },
        },
        {
          sdt: {
            properties: { alias: "Agree", tag: "agree", checkbox: { checked: true } },
          },
        },

        { paragraph: { children: [""] } },

        // 4. Custom checkbox symbols (✔ / ✕ instead of ☒ / ☐)
        {
          paragraph: {
            children: [{ bold: true, text: "4. Custom symbols", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          sdt: {
            properties: {
              alias: "Feature",
              tag: "feature",
              checkbox: {
                checked: true,
                checkedState: { val: "2714", font: "MS Gothic" }, // ✔
                uncheckedState: { val: "2715", font: "MS Gothic" }, // ✕
              },
            },
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
