// Demo: SmartArt - all layout types with various styles and colors
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, SmartArtRun } from "docx-plus";

const heading = (text: string) =>
    new Paragraph({
        children: [new TextRun({ text, bold: true, size: 32 })],
        spacing: { after: 200 },
    });

const subheading = (text: string) =>
    new Paragraph({
        children: [new TextRun({ text, bold: true, size: 24, color: "4472C4" })],
        spacing: { before: 300, after: 100 },
    });

const note = () =>
    new Paragraph({
        children: [
            new TextRun({
                text: "Note: Open this document in Microsoft Word to see the SmartArt rendered.",
                italics: true,
                color: "888888",
                size: 18,
            }),
        ],
        spacing: { before: 600 },
    });

const doc = new Document({
    sections: [
        {
            children: [
                heading("SmartArt Layouts Demo"),

                // --- 1. List Layouts ---
                subheading("List Layouts"),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    {
                                        text: "Item A",
                                        children: [{ text: "Sub A1" }, { text: "Sub A2" }],
                                    },
                                    { text: "Item B", children: [{ text: "Sub B1" }] },
                                    { text: "Item C" },
                                ],
                            },
                            transformation: { width: 450, height: 250 },
                            layout: "default",
                            style: "simple1",
                            color: "accent1_2",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [{ text: "Step 1" }, { text: "Step 2" }, { text: "Step 3" }],
                            },
                            transformation: { width: 450, height: 250 },
                            layout: "list1",
                            style: "simple2",
                            color: "accent2_2",
                        }),
                    ],
                }),

                // --- 2. Process Layouts ---
                subheading("Process Layouts"),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Plan" },
                                    { text: "Design" },
                                    { text: "Build" },
                                    { text: "Test" },
                                    { text: "Deploy" },
                                ],
                            },
                            transformation: { width: 550, height: 250 },
                            layout: "process1",
                            style: "moderate1",
                            color: "accent3_2",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Phase 1" },
                                    { text: "Phase 2" },
                                    { text: "Phase 3" },
                                    { text: "Phase 4" },
                                ],
                            },
                            transformation: { width: 550, height: 250 },
                            layout: "chevron1",
                            style: "professional1",
                            color: "accent4_2",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Start" },
                                    { text: "Process" },
                                    { text: "Review" },
                                    { text: "Complete" },
                                ],
                            },
                            transformation: { width: 550, height: 250 },
                            layout: "arrow1",
                            style: "polished1",
                            color: "accent5_2",
                        }),
                    ],
                }),

                // --- 3. Cycle Layouts ---
                subheading("Cycle Layouts"),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Research" },
                                    { text: "Design" },
                                    { text: "Develop" },
                                    { text: "Test" },
                                ],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "cycle1",
                            style: "cartoon1",
                            color: "colorful1",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Plan" },
                                    { text: "Do" },
                                    { text: "Check" },
                                    { text: "Act" },
                                    { text: "Improve" },
                                ],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "cycle2",
                            style: "powdery1",
                            color: "colorful2",
                        }),
                    ],
                }),

                // --- 4. Hierarchy Layouts ---
                subheading("Hierarchy Layouts"),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    {
                                        text: "CEO",
                                        children: [
                                            {
                                                text: "VP Engineering",
                                                children: [
                                                    { text: "Frontend" },
                                                    { text: "Backend" },
                                                ],
                                            },
                                            {
                                                text: "VP Marketing",
                                                children: [{ text: "Brand" }, { text: "Growth" }],
                                            },
                                        ],
                                    },
                                ],
                            },
                            transformation: { width: 500, height: 350 },
                            layout: "hierarchy1",
                            style: "moderate2",
                            color: "accent1_2",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    {
                                        text: "Root",
                                        children: [
                                            {
                                                text: "Branch A",
                                                children: [
                                                    { text: "Leaf A1" },
                                                    { text: "Leaf A2" },
                                                ],
                                            },
                                            { text: "Branch B", children: [{ text: "Leaf B1" }] },
                                            { text: "Branch C" },
                                        ],
                                    },
                                ],
                            },
                            transformation: { width: 500, height: 350 },
                            layout: "hierarchy2",
                            style: "polished2",
                            color: "colorful3",
                        }),
                    ],
                }),

                // --- 5. Relationship Layouts ---
                subheading("Relationship / Other Layouts"),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [{ text: "Set A" }, { text: "Set B" }, { text: "Set C" }],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "venn1",
                            style: "burnt1",
                            color: "accent6_2",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Leads" },
                                    { text: "Prospects" },
                                    { text: "Qualified" },
                                    { text: "Negotiation" },
                                    { text: "Closed" },
                                ],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "funnel1",
                            style: "professional2",
                            color: "dark1",
                        }),
                    ],
                }),

                // --- 6. Pyramid Layouts ---
                subheading("Pyramid Layouts"),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Vision" },
                                    { text: "Strategy" },
                                    { text: "Tactics" },
                                    { text: "Operations" },
                                ],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "pyramid1",
                            style: "moderate3",
                            color: "accent2_2",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "All" },
                                    { text: "Most" },
                                    { text: "Some" },
                                    { text: "Few" },
                                ],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "pyramid2",
                            style: "cartoon2",
                            color: "colorful4",
                        }),
                    ],
                }),

                // --- 7. Matrix / Radial / Gear / Balance ---
                subheading("Matrix / Radial / Other Layouts"),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Q1: Plan" },
                                    { text: "Q2: Execute" },
                                    { text: "Q3: Review" },
                                    { text: "Q4: Deliver" },
                                ],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "matrix1",
                            style: "professional3",
                            color: "primary1",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    { text: "Core" },
                                    { text: "Team" },
                                    { text: "Tools" },
                                    { text: "Process" },
                                ],
                            },
                            transformation: { width: 350, height: 300 },
                            layout: "radial1",
                            style: "burnt2",
                            color: "gray1",
                        }),
                    ],
                }),

                note(),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
