// Table of contents

import { writeFileSync } from "node:fs";

import { HeadingLevel, generateDocument } from "@office-open/docx";

// WordprocessingML docs for TableOfContents can be found here:
// http://officeopenxml.com/WPtableOfContents.php

// Let's define the properties for generate a TOC for heading 1-5 and MySpectacularStyle,
// Making the entries be hyperlinks for the paragraph
const buffer = await generateDocument({
  features: {
    updateFields: true,
  },
  sections: [
    {
      children: [
        {
          toc: {
            alias: "Summary",
            headingStyleRange: "1-5",
            hyperlink: true,
            stylesWithLevels: [{ styleName: "MySpectacularStyle", level: 1 }],
          },
        },
        {
          paragraph: {
            children: [{ bookmarkStart: { id: 0, name: "anchorForHeader1" } }, { bookmarkEnd: 0 }],
            heading: HeadingLevel.HEADING_1,
            pageBreakBefore: true,
            text: "Header #1",
          },
        },
        { paragraph: "I'm a little text very nicely written.'" },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            pageBreakBefore: true,
            text: "Header #2",
          },
        },
        { paragraph: "I'm a other text very nicely written.'" },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "Header #2.1",
          },
        },
        { paragraph: "I'm a another text very nicely written.'" },
        {
          paragraph: {
            pageBreakBefore: true,
            style: "MySpectacularStyle",
            text: "My Spectacular Style #1",
          },
        },
      ],
    },
  ],
  styles: {
    paragraphStyles: [
      {
        basedOn: "Heading1",
        id: "MySpectacularStyle",
        name: "My Spectacular Style",
        next: "Heading1",
        quickFormat: true,
        run: {
          color: "990000",
          italic: true,
        },
      },
      {
        basedOn: "Heading2",
        id: "TOC2",
        name: "TOC 2",
        paragraph: {
          indent: {
            left: 240,
          },
        },
        quickFormat: true,
      },
    ],
  },
});
writeFileSync("My Document.docx", buffer);
