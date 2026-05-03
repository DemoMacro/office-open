// Table of contents

import * as fs from "fs";

import {
    Bookmark,
    File,
    HeadingLevel,
    Packer,
    Paragraph,
    StyleLevel,
    TableOfContents,
} from "@office-open/docx";

// WordprocessingML docs for TableOfContents can be found here:
// http://officeopenxml.com/WPtableOfContents.php

// Let's define the properties for generate a TOC for heading 1-5 and MySpectacularStyle,
// Making the entries be hyperlinks for the paragraph
const doc = new File({
    features: {
        updateFields: true,
    },
    sections: [
        {
            children: [
                new TableOfContents("Summary", {
                    cachedEntries: [
                        {
                            title: "Header #1",
                            level: 1,
                            page: 1,
                            href: "anchorForHeader1",
                        },
                        {
                            title: "Header #2",
                            level: 1,
                            page: 2,
                        },
                        {
                            title: "Header #2.1",
                            level: 2,
                        },
                        {
                            title: "My Spectacular Style #1",
                            level: 1,
                            page: 3,
                        },
                    ],
                    headingStyleRange: "1-5",
                    hyperlink: true,
                    stylesWithLevels: [new StyleLevel("MySpectacularStyle", 1)],
                }),
                new Paragraph({
                    children: [
                        new Bookmark({
                            id: "anchorForHeader1",
                            children: [],
                        }),
                    ],
                    heading: HeadingLevel.HEADING_1,
                    pageBreakBefore: true,
                    text: "Header #1",
                }),
                new Paragraph("I'm a little text very nicely written.'"),
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    pageBreakBefore: true,
                    text: "Header #2",
                }),
                new Paragraph("I'm a other text very nicely written.'"),
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "Header #2.1",
                }),
                new Paragraph("I'm a another text very nicely written.'"),
                new Paragraph({
                    pageBreakBefore: true,
                    style: "MySpectacularStyle",
                    text: "My Spectacular Style #1",
                }),
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
                    italics: true,
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

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
