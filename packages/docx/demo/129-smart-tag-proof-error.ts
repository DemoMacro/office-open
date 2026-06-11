// Smart tags and proof errors: SmartTagRun wraps recognized text,
// ProofError marks spelling/grammar error ranges.

import { writeFileSync } from "node:fs";

import { ProofErrorType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ bold: true, text: "Smart Tags and Proof Errors", size: 16 }],
          },
        },

        // Smart tag wrapping a date
        {
          paragraph: {
            children: [
              "Today is ",
              {
                smartTag: {
                  uri: "urn:schemas-microsoft-com:office:smarttags",
                  element: "date",
                  properties: [
                    {
                      uri: "urn:schemas-microsoft-com:office:smarttags",
                      name: "month",
                      val: "June",
                    },
                  ],
                  children: ["June 4, 2026"],
                },
              },
              ".",
            ],
          },
        },

        // Proof error marking a spelling mistake
        {
          paragraph: {
            children: [
              "This paragraph demonstrates ",
              { proofErr: ProofErrorType.SPELL_START },
              "teh",
              { proofErr: ProofErrorType.SPELL_END },
              " proof error markers.",
            ],
          },
        },

        // Grammar error range
        {
          paragraph: {
            children: [
              "Grammar error example: ",
              { proofErr: ProofErrorType.GRAM_START },
              "Their going to the store.",
              { proofErr: ProofErrorType.GRAM_END },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
