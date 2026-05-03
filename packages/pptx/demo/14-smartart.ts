import { Presentation, Packer, Shape, Paragraph, Run, SmartArtFrame } from "../src";

const pres = new Presentation({
    title: "SmartArt Demo",
    creator: "Demo",
    slides: [
        // --- 1. List Layouts ---
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [new Run({ text: "List Layouts", fontSize: 32, bold: true })],
                        }),
                    ],
                }),
                new SmartArtFrame({
                    x: 50,
                    y: 120,
                    width: 400,
                    height: 300,
                    nodes: [
                        { text: "Item A", children: [{ text: "Sub A1" }, { text: "Sub A2" }] },
                        { text: "Item B", children: [{ text: "Sub B1" }] },
                        { text: "Item C" },
                    ],
                    layout: "default",
                    style: "simple1",
                    color: "accent1_2",
                }),
                new SmartArtFrame({
                    x: 480,
                    y: 120,
                    width: 400,
                    height: 300,
                    nodes: [{ text: "Step 1" }, { text: "Step 2" }, { text: "Step 3" }],
                    layout: "list1",
                    style: "simple2",
                    color: "accent2_2",
                }),
            ],
        },

        // --- 2. Process Layouts ---
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({ text: "Process Layouts", fontSize: 32, bold: true }),
                            ],
                        }),
                    ],
                }),
                new SmartArtFrame({
                    x: 50,
                    y: 120,
                    width: 850,
                    height: 300,
                    nodes: [
                        { text: "Plan" },
                        { text: "Design" },
                        { text: "Build" },
                        { text: "Test" },
                        { text: "Deploy" },
                    ],
                    layout: "process1",
                    style: "moderate1",
                    color: "accent3_2",
                }),
            ],
        },
        {
            children: [
                new SmartArtFrame({
                    x: 50,
                    y: 50,
                    width: 850,
                    height: 350,
                    nodes: [
                        { text: "Phase 1" },
                        { text: "Phase 2" },
                        { text: "Phase 3" },
                        { text: "Phase 4" },
                    ],
                    layout: "chevron1",
                    style: "professional1",
                    color: "accent4_2",
                }),
            ],
        },
        {
            children: [
                new SmartArtFrame({
                    x: 50,
                    y: 50,
                    width: 850,
                    height: 350,
                    nodes: [
                        { text: "Start" },
                        { text: "Process" },
                        { text: "Review" },
                        { text: "Complete" },
                    ],
                    layout: "arrow1",
                    style: "polished1",
                    color: "accent5_2",
                }),
            ],
        },

        // --- 3. Cycle Layouts ---
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({ text: "Cycle Layouts", fontSize: 32, bold: true }),
                            ],
                        }),
                    ],
                }),
                new SmartArtFrame({
                    x: 50,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [
                        { text: "Research" },
                        { text: "Design" },
                        { text: "Develop" },
                        { text: "Test" },
                    ],
                    layout: "cycle1",
                    style: "cartoon1",
                    color: "colorful1",
                }),
                new SmartArtFrame({
                    x: 480,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [
                        { text: "Plan" },
                        { text: "Do" },
                        { text: "Check" },
                        { text: "Act" },
                        { text: "Improve" },
                    ],
                    layout: "cycle2",
                    style: "powdery1",
                    color: "colorful2",
                }),
            ],
        },

        // --- 4. Hierarchy Layouts ---
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({ text: "Hierarchy Layouts", fontSize: 32, bold: true }),
                            ],
                        }),
                    ],
                }),
                new SmartArtFrame({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 400,
                    nodes: [
                        {
                            text: "CEO",
                            children: [
                                {
                                    text: "VP Engineering",
                                    children: [{ text: "Frontend" }, { text: "Backend" }],
                                },
                                {
                                    text: "VP Marketing",
                                    children: [{ text: "Brand" }, { text: "Growth" }],
                                },
                            ],
                        },
                    ],
                    layout: "hierarchy1",
                    style: "moderate2",
                    color: "accent1_2",
                }),
            ],
        },
        {
            children: [
                new SmartArtFrame({
                    x: 50,
                    y: 50,
                    width: 600,
                    height: 400,
                    nodes: [
                        {
                            text: "Root",
                            children: [
                                {
                                    text: "Branch A",
                                    children: [{ text: "Leaf A1" }, { text: "Leaf A2" }],
                                },
                                { text: "Branch B", children: [{ text: "Leaf B1" }] },
                                { text: "Branch C" },
                            ],
                        },
                    ],
                    layout: "hierarchy2",
                    style: "polished2",
                    color: "colorful3",
                }),
            ],
        },

        // --- 5. Relationship Layouts ---
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({
                                    text: "Relationship / Other Layouts",
                                    fontSize: 32,
                                    bold: true,
                                }),
                            ],
                        }),
                    ],
                }),
                new SmartArtFrame({
                    x: 50,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [{ text: "Set A" }, { text: "Set B" }, { text: "Set C" }],
                    layout: "venn1",
                    style: "burnt1",
                    color: "accent6_2",
                }),
                new SmartArtFrame({
                    x: 480,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [
                        { text: "Leads" },
                        { text: "Prospects" },
                        { text: "Qualified" },
                        { text: "Negotiation" },
                        { text: "Closed" },
                    ],
                    layout: "funnel1",
                    style: "professional2",
                    color: "dark1",
                }),
            ],
        },

        // --- 6. Pyramid Layouts ---
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({ text: "Pyramid Layouts", fontSize: 32, bold: true }),
                            ],
                        }),
                    ],
                }),
                new SmartArtFrame({
                    x: 50,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [
                        { text: "Vision" },
                        { text: "Strategy" },
                        { text: "Tactics" },
                        { text: "Operations" },
                    ],
                    layout: "pyramid1",
                    style: "moderate3",
                    color: "accent2_2",
                }),
                new SmartArtFrame({
                    x: 480,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [{ text: "All" }, { text: "Most" }, { text: "Some" }, { text: "Few" }],
                    layout: "pyramid2",
                    style: "cartoon2",
                    color: "colorful4",
                }),
            ],
        },

        // --- 7. Matrix / Radial / Other Layouts ---
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({
                                    text: "Matrix / Radial / Other Layouts",
                                    fontSize: 32,
                                    bold: true,
                                }),
                            ],
                        }),
                    ],
                }),
                new SmartArtFrame({
                    x: 50,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [
                        { text: "Q1: Plan" },
                        { text: "Q2: Execute" },
                        { text: "Q3: Review" },
                        { text: "Q4: Deliver" },
                    ],
                    layout: "matrix1",
                    style: "professional3",
                    color: "primary1",
                }),
                new SmartArtFrame({
                    x: 480,
                    y: 120,
                    width: 400,
                    height: 350,
                    nodes: [
                        { text: "Core" },
                        { text: "Team" },
                        { text: "Tools" },
                        { text: "Process" },
                    ],
                    layout: "radial1",
                    style: "burnt2",
                    color: "gray1",
                }),
            ],
        },
        {
            children: [
                new SmartArtFrame({
                    x: 50,
                    y: 50,
                    width: 400,
                    height: 350,
                    nodes: [{ text: "Input" }, { text: "Process" }, { text: "Output" }],
                    layout: "balance1",
                    style: "powdery2",
                    color: "primary2",
                }),
                new SmartArtFrame({
                    x: 480,
                    y: 50,
                    width: 400,
                    height: 350,
                    nodes: [
                        { text: "Part A" },
                        { text: "Part B" },
                        { text: "Part C" },
                        { text: "Part D" },
                    ],
                    layout: "gear1",
                    style: "polished3",
                    color: "accent3_2",
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
import { writeFileSync } from "fs";
writeFileSync("My Presentation.pptx", buffer);
