// Demo: Permission Ranges (Document Protection - Editable Regions)
// Demonstrates PermStart and PermEnd for defining editable regions.
// Note: documentProtection must be enabled in features for permissions to take effect.

import * as fs from "fs";

import { Document, EditGroupType, Packer, Paragraph, PermEnd, PermStart, TextRun } from "docx-plus";

const doc = new Document({
    features: {
        trackRevisions: true,
        documentProtection: {
            edit: "readOnly",
            hashValue:
                "OnpmSDysb0NPnL/m//Xu3LuX3K1Tg0hxqTZmTIPOi86G/mlRwOXhcd4DOppG/LwyZi9BxWrn+N7Os3DQBbjuQA==",
            saltValue: "E3ilFmh48B4vQPAQpIuZXA==",
            spinCount: 100000,
            algorithmName: "SHA-512",
        },
    },
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({ text: "Permission Ranges Demo", bold: true, size: 32 }),
                    ],
                    spacing: { after: 400 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Password: 123 | Protection: Read-only (with editable ranges)"),
                    ],
                }),

                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "1. Everyone Editable Range", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [new TextRun("This text is read-only (protected).")],
                }),
                new Paragraph({
                    children: [
                        new PermStart({ id: 0, edGroup: EditGroupType.EVERYONE }),
                        new TextRun("Everyone can edit this text."),
                        new PermEnd(0),
                    ],
                }),
                new Paragraph({
                    children: [new TextRun("This text is read-only again.")],
                }),

                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "2. Single User Editable Range",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new PermStart({ id: 1, ed: "john@example.com" }),
                        new TextRun("Only john@example.com can edit this text."),
                        new PermEnd(1),
                    ],
                }),

                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "3. Editors Group Range", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new PermStart({ id: 2, edGroup: EditGroupType.EDITORS }),
                        new TextRun("Only editors can modify this content."),
                        new PermEnd(2),
                    ],
                }),

                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "4. Current User Only Range", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new PermStart({ id: 3, edGroup: EditGroupType.CURRENT }),
                        new TextRun("Only the current user can edit this."),
                        new PermEnd(3),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
