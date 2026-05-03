import * as fs from "fs";

import { Presentation, Packer, Shape } from "@office-open/pptx";

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
                    text: "Slide 1 - Default footer",
                    fill: "4472C4",
                }),
            ],
            headerFooter: { slideNumber: true, dateTime: true, footer: "Confidential" },
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Slide 2 - No date",
                    fill: "ED7D31",
                }),
            ],
            headerFooter: { slideNumber: true, dateTime: false, footer: "My Presentation" },
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Slide 3 - Only slide number",
                    fill: "70AD47",
                }),
            ],
            headerFooter: { slideNumber: true, dateTime: false, footer: false },
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Slide 4 - No header/footer",
                    fill: "FFC000",
                }),
            ],
            headerFooter: { slideNumber: false, dateTime: false, footer: false },
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
