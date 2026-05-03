// Patch a document with patches

import * as fs from "fs";

import { PatchType, TextRun, patchDocument } from "@office-open/docx";

const doc = await patchDocument({
    data: fs.readFileSync("demo/assets/field-trip.docx"),
    outputType: "nodebuffer",
    patches: {
        address: {
            children: [new TextRun({ text: "blah blah" })],
            type: PatchType.PARAGRAPH,
        },

        city: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        email_address: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        first_name: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        ft_dates: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        grade: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        last_name: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        phone: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        school_name: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        state: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },

        todays_date: {
            children: [new TextRun({ text: new Date().toLocaleDateString() })],
            type: PatchType.PARAGRAPH,
        },

        zip: {
            children: [new TextRun({ text: "test" })],
            type: PatchType.PARAGRAPH,
        },
    },
});
fs.writeFileSync("My Document.docx", doc);
