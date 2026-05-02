import * as fs from "fs";

import { Presentation, Packer, Shape, SolidFill } from "../src";

const pres = new Presentation({
    title: "Notes Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Slide 1 - Introduction",
                    fill: new SolidFill("4472C4"),
                }),
                new Shape({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 300,
                    text: "Welcome to the presentation!",
                    fill: new SolidFill("D9E2F3"),
                }),
            ],
            notes: "These are the speaker notes for the introduction slide. Remember to greet the audience.",
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Slide 2 - Key Points",
                    fill: new SolidFill("ED7D31"),
                }),
                new Shape({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 300,
                    text: "Point 1: Architecture\nPoint 2: Implementation\nPoint 3: Testing",
                    fill: new SolidFill("FBE5D6"),
                }),
            ],
            notes: "Key talking points:\n- Architecture follows SOLID principles\n- Implementation uses OOXML spec\n- Testing covers all chart types",
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Slide 3 - No Notes",
                    fill: new SolidFill("70AD47"),
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
