import * as fs from "fs";

import { Background, Presentation, Packer, Shape, SolidFill, TableFrame } from "../src";

const pres = new Presentation({
    title: "Phase 3 Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Table Demo",
                    fill: new SolidFill("4472C4"),
                }),
                new TableFrame({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 250,
                    rows: [
                        {
                            cells: [
                                {
                                    text: "Name",
                                    fill: new SolidFill("4472C4"),
                                },
                                { text: "Age" },
                                { text: "City" },
                            ],
                        },
                        {
                            cells: [{ text: "Alice" }, { text: "30" }, { text: "Beijing" }],
                        },
                        {
                            cells: [{ text: "Bob" }, { text: "25" }, { text: "Shanghai" }],
                        },
                        {
                            cells: [{ text: "Charlie" }, { text: "35" }, { text: "Shenzhen" }],
                        },
                    ],
                    columnWidths: [2000000, 1500000, 2500000],
                    firstRow: true,
                    bandRow: true,
                }),
            ],
        },
        {
            background: new Background({ fill: new SolidFill("F5F5F5") }),
            children: [
                new TableFrame({
                    x: 100,
                    y: 80,
                    width: 500,
                    height: 200,
                    rows: [
                        {
                            cells: [
                                {
                                    text: "Header 1",
                                    fill: new SolidFill("ED7D31"),
                                },
                                { text: "Header 2", fill: new SolidFill("ED7D31") },
                            ],
                        },
                        {
                            cells: [{ text: "A" }, { text: "B" }],
                        },
                        {
                            cells: [{ text: "C" }, { text: "D" }],
                        },
                    ],
                    columnWidths: [2500000, 2500000],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
