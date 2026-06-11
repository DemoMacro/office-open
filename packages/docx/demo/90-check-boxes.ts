// NOTE: The CheckBox class (w14:checkbox SDT content control) has been removed.
// This demo now renders static checkbox symbols using symbolRun instead of
// interactive checkboxes. For form field checkboxes, see form-field.ts.
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              { break: 1 },
              // Unchecked (default symbol: ☐ 2610, MS Gothic)
              { symbolRun: { char: "2610", symbolfont: "MS Gothic" } },
              { break: 1 },
              // Checked (default symbol: ☒ 2612, MS Gothic)
              { symbolRun: { char: "2612", symbolfont: "MS Gothic" } },
              { break: 1 },
              // Checked with custom checkedState value 2611 (☑)
              { symbolRun: { char: "2611", symbolfont: "MS Gothic" } },
              { break: 1 },
              // Checked with custom font + value
              { symbolRun: { char: "2611", symbolfont: "MS Gothic" } },
              { break: 1 },
              // Checked with custom checked + unchecked state symbols
              { symbolRun: { char: "2611", symbolfont: "MS Gothic" } },
              { break: 1 },
              // Checked with same symbols as above
              { symbolRun: { char: "2611", symbolfont: "MS Gothic" } },
              { break: 1, text: "Are you ok?" },
              // Checked checkbox with alias
              { symbolRun: { char: "2612", symbolfont: "MS Gothic" } },
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
