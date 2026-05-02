import * as fs from "fs";

import { Presentation, Shape, Packer, SolidFill } from "../src";

const pres = new Presentation({
    title: "My Presentation",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 400,
                    height: 200,
                    text: "Hello World",
                    geometry: "rect",
                    fill: new SolidFill("4472C4"),
                }),
                new Shape({
                    x: 200,
                    y: 350,
                    width: 500,
                    height: 100,
                    text: "Second shape",
                }),
            ],
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 50,
                    width: 860,
                    height: 440,
                    text: "Slide 2 - Full Width",
                    geometry: "rect",
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
