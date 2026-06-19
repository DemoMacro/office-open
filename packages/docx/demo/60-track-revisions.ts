// Track Revisions aka. "Track Changes"

import { writeFileSync } from "node:fs";

import { AlignmentType, PageNumber, ShadingType, generateDocument } from "@office-open/docx";

/*
    For reference, see
    - https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.insertedrun
    - https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.deletedrun

    The setting `features: { trackRevisions: true }` adds an element `<w:trackRevisions />` to the `settings.xml` file.
    This specifies that the application shall track *new* revisions made to the existing document.
    See also https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.trackrevisions

    Note that this setting enables to track *new changes* after teh file is generated, so this example will still
    show inserted and deleted text runs when you remove it.
*/

const paragraph = {
  paragraph: {
    children: [
      "This is a simple demo ",
      {
        text: "on how to ",
      },
      {
        insertion: {
          author: "Firstname Lastname",
          date: "2020-10-06T09:00:00Z",
          id: 0,
          children: [{ text: "mark a text as an insertion " }],
        },
      },
      {
        deletion: {
          author: "Firstname Lastname",
          date: "2020-10-06T09:00:00Z",
          id: 1,
          children: [{ text: "or a deletion." }],
        },
      },
    ],
  },
};

const buffer = await generateDocument({
  features: {
    trackRevisions: true,
  },
  footnotes: {
    1: {
      children: [
        {
          children: [
            "This is a footnote",
            {
              deletion: {
                author: "Firstname Lastname",
                date: "2020-10-06T09:05:00Z",
                id: 0,
                children: [{ text: " with some extra text which was deleted" }],
              },
            },
            {
              insertion: {
                author: "Firstname Lastname",
                date: "2020-10-06T09:05:00Z",
                id: 1,
                children: [{ text: " and new content" }],
              },
            },
          ],
        },
      ],
    },
  },
  sections: [
    {
      children: [
        paragraph,
        {
          paragraph: {
            children: [
              "This is a demo ",
              {
                deletion: {
                  id: 2,
                  author: "Firstname Lastname",
                  date: "2020-10-06T09:00:00Z",
                  children: [
                    {
                      break: 1,
                      text: "in order",
                      color: "ff0000",
                      bold: true,
                      size: 12,
                      font: {
                        name: "Garamond",
                      },
                      shading: {
                        type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                        color: "00FFFF",
                        fill: "FF0000",
                      },
                    },
                  ],
                },
              },
              {
                insertion: {
                  id: 3,
                  author: "Firstname Lastname",
                  date: "2020-10-06T09:05:00Z",
                  children: [
                    {
                      text: "to show how to ",
                      bold: false,
                    },
                  ],
                },
              },
              {
                bold: true,
                children: [{ tab: true }, "use Inserted and Deleted TextRuns."],
              },
              { footnoteReference: 1 },
              {
                bold: true,
                text: "And some style changes",
                revision: {
                  id: 4,
                  author: "Firstname Lastname",
                  date: "2020-10-06T09:05:00Z",
                  bold: false,
                },
              },
            ],
          },
        },
      ],
      footers: {
        default: [
          {
            paragraph: {
              alignment: AlignmentType.CENTER,
              children: [
                "Awesome LLC",
                {
                  children: ["Page Number: ", PageNumber.CURRENT],
                },
                {
                  deletion: {
                    children: [" to ", PageNumber.TOTAL_PAGES],
                    id: 4,
                    author: "Firstname Lastname",
                    date: "2020-10-06T09:05:00Z",
                  },
                },
                {
                  insertion: {
                    children: [{ text: " from ", bold: true }, PageNumber.TOTAL_PAGES],
                    id: 5,
                    author: "Firstname Lastname",
                    date: "2020-10-06T09:05:00Z",
                  },
                },
              ],
            },
          },
        ],
      },
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
