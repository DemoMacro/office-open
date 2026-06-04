// Smart tags and proof errors: SmartTagRun wraps recognized text,
// ProofError marks spelling/grammar error ranges.

import * as fs from "fs";

import {
  Document,
  Packer,
  Paragraph,
  ProofError,
  ProofErrorType,
  SmartTagRun,
  TextRun,
} from "@office-open/docx";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [new TextRun({ bold: true, text: "Smart Tags and Proof Errors", size: 32 })],
        }),

        // Smart tag wrapping a date
        new Paragraph({
          children: [
            new TextRun("Today is "),
            new SmartTagRun({
              uri: "urn:schemas-microsoft-com:office:smarttags",
              element: "date",
              properties: [
                {
                  uri: "urn:schemas-microsoft-com:office:smarttags",
                  name: "month",
                  val: "June",
                },
              ],
              children: [new TextRun("June 4, 2026")],
            }),
            new TextRun("."),
          ],
        }),

        // Proof error marking a spelling mistake
        new Paragraph({
          children: [
            new TextRun("This paragraph demonstrates "),
            new ProofError(ProofErrorType.SPELL_START),
            new TextRun("teh"),
            new ProofError(ProofErrorType.SPELL_END),
            new TextRun(" proof error markers."),
          ],
        }),

        // Grammar error range
        new Paragraph({
          children: [
            new TextRun("Grammar error example: "),
            new ProofError(ProofErrorType.GRAM_START),
            new TextRun("Their going to the store."),
            new ProofError(ProofErrorType.GRAM_END),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
