// Patch a document with patches

import * as fs from "fs";

import { PatchType, TextRun, patchDocument } from "docx";

patchDocument({
    data: fs.readFileSync("demo/assets/simple-template-2.docx"),
    outputType: "nodebuffer",
    patches: {
        name: {
            children: [new TextRun("Max")],
            type: PatchType.PARAGRAPH,
        },
    },
}).then((doc) => {
    fs.writeFileSync("My Document.docx", doc);
});
