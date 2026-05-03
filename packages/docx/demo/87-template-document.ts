// Patch a document with patches

import * as fs from "fs";

import { PatchType, TextRun, patchDocument } from "@office-open/docx";

const doc = await patchDocument({
    data: fs.readFileSync("demo/assets/simple-template-2.docx"),
    outputType: "nodebuffer",
    patches: {
        name: {
            children: [new TextRun("Max")],
            type: PatchType.PARAGRAPH,
        },
    },
});
fs.writeFileSync("My Document.docx", doc);
