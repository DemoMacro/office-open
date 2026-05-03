// Patch a document with patches

import * as fs from "fs";

import type { IPatch } from "@office-open/docx";
import { patchDocument, PatchType, TextRun } from "@office-open/docx";

export const font = "Trebuchet MS";
export const getPatches = (fields: Record<string, string>) => {
    const patches: Record<string, IPatch> = {};

    for (const field in fields) {
        patches[field] = {
            children: [new TextRun({ font, text: fields[field] })],
            type: PatchType.PARAGRAPH,
        };
    }

    return patches;
};

const patches = getPatches({
    "first-name": "John",
    salutation: "Mr.",
});

const doc = await patchDocument({
    data: fs.readFileSync("demo/assets/simple-template-3.docx"),
    keepOriginalStyles: true,
    outputType: "nodebuffer",
    patches,
});
fs.writeFileSync("My Document.docx", doc);
