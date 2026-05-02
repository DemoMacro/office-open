import * as fs from "fs";

import { Presentation, Packer, Shape, SolidFill, Paragraph, Run } from "../src";

const pres = new Presentation({
    title: "Header Footer Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Slide 1",
                    fill: new SolidFill("4472C4"),
                }),
            ],
            headerFooter: {
                slideNumber: true,
                dateTime: true,
                footer: "Confidential",
            },
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Slide 2 - No Footer",
                    fill: new SolidFill("ED7D31"),
                }),
            ],
            headerFooter: {
                slideNumber: true,
                dateTime: false,
            },
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Slide 3 - Custom Footer",
                    fill: new SolidFill("70AD47"),
                }),
            ],
            headerFooter: {
                slideNumber: true,
                footer: "My Presentation",
            },
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Slide 4 - No Header/Footer",
                    fill: new SolidFill("FFC000"),
                }),
            ],
            headerFooter: {
                slideNumber: false,
                dateTime: false,
                footer: undefined,
            },
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
