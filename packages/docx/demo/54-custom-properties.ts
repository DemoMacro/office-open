// Custom Properties
// Custom properties are incredibly useful if you want to be able to apply quick parts or custom cover pages
// To the document in Word after the document has been generated. Standard properties (such as creator, title
// And subject) cover typical use cases, but sometimes custom properties are required.

import * as fs from "fs";

import { Document, Packer } from "@office-open/docx";

const doc = new Document(
    // Standard properties
    {
        creator: "Creator",
        customProperties: [
            { name: "Subtitle", value: "Subtitle" },
            { name: "Address", value: "Address" },
        ],
        description: "Description",
        sections: [],
        subject: "Subject",
        title: "Title",
    },
);

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
