// Custom styles using JavaScript configuration

import * as fs from "fs";

import {
    Document,
    HeadingLevel,
    Packer,
    Paragraph,
    UnderlineType,
    convertInchesToTwip,
} from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    style: "myWonkyStyle",
                    text: "Hello",
                }),
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "World",
                }),
            ],
        },
    ],
    styles: {
        paragraphStyles: [
            {
                basedOn: "Normal",
                id: "myWonkyStyle",
                name: "My Wonky Style",
                next: "Normal",
                paragraph: {
                    indent: {
                        left: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                    },
                },
                run: {
                    color: "990000",
                    italics: true,
                },
            },
            {
                basedOn: "Normal",
                id: "Heading2",
                name: "Heading 2",
                next: "Normal",
                paragraph: {
                    spacing: {
                        after: 120,
                        before: 240,
                    },
                },
                quickFormat: true,
                run: {
                    bold: true,
                    size: 26,
                    underline: {
                        color: "FF0000",
                        type: UnderlineType.DOUBLE,
                    },
                },
            },
        ],
    },
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
