// This demo shows how to create bookmarks then link to them with internal hyperlinks

import { writeFileSync } from "node:fs";

import { HeadingLevel, generateDocument } from "@office-open/docx";

const LOREM_IPSUM =
  /* Cspell:disable-next-line */
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam mi velit, convallis convallis scelerisque nec, faucibus nec leo. Phasellus at posuere mauris, tempus dignissim velit. Integer et tortor dolor. Duis auctor efficitur mattis. Vivamus ut metus accumsan tellus auctor sollicitudin venenatis et nibh. Cras quis massa ac metus fringilla venenatis. Proin rutrum mauris purus, ut suscipit magna consectetur id. Integer consectetur sollicitudin ante, vitae faucibus neque efficitur in. Praesent ultricies nibh lectus. Mauris pharetra id odio eget iaculis. Duis dictum, risus id pellentesque rutrum, lorem quam malesuada massa, quis ullamcorper turpis urna a diam. Cras vulputate metus vel massa porta ullamcorper. Etiam porta condimentum nulla nec tristique. Sed nulla urna, pharetra non tortor sed, sollicitudin molestie diam. Maecenas enim leo, feugiat eget vehicula id, sollicitudin vitae ante.";

const buffer = await generateDocument({
  creator: "Clippy",
  description: "A brief example of using docx with bookmarks and internal hyperlinks",
  sections: [
    {
      children: [
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            children: [
              { bookmarkStart: { id: 0, name: "myAnchorId" } },
              "Lorem Ipsum",
              { bookmarkEnd: 0 },
            ],
          },
        },
        { paragraph: "\n" },
        { paragraph: LOREM_IPSUM },
        {
          paragraph: {
            children: [{ pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: [
              {
                hyperlink: {
                  children: [
                    {
                      text: "Styled",
                      bold: true,
                      style: "Hyperlink",
                    },
                    {
                      text: " Anchor Text",
                      style: "Hyperlink",
                    },
                  ],
                  anchor: "myAnchorId",
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "The bookmark can be seen on page ",
              { pageReference: { bookmarkId: "myAnchorId" } },
            ],
          },
        },
      ],
      footers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  hyperlink: {
                    children: [
                      {
                        text: "Click here!",
                        style: "Hyperlink",
                      },
                    ],
                    anchor: "myAnchorId",
                  },
                },
              ],
            },
          },
        ],
      },
    },
  ],
  title: "Sample Document",
});
writeFileSync("My Document.docx", buffer);
