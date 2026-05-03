// This example shows 3 styles using XML styles

import * as fs from "fs";

import { Document, HeadingLevel, Packer, Paragraph } from "@office-open/docx";

const styles = fs.readFileSync("./demo/assets/custom-styles.xml", "utf8");
const doc = new Document({
    externalStyles: styles,
    sections: [
        {
            children: [
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "Cool Heading Text",
                }),
                new Paragraph({
                    style: "MyFancyStyle",
                    text: 'This is a custom named style from the template "MyFancyStyle"',
                }),
                new Paragraph("Some normal text"),
                new Paragraph({
                    style: "MyFancyStyle",
                    text: "MyFancyStyle again",
                }),
            ],
        },
    ],
    title: "Title",
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
