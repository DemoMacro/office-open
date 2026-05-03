import * as fs from "fs";

import { Background, ChartFrame, Presentation, Packer, Shape, SolidFill } from "@office-open/pptx";

const pres = new Presentation({
    title: "Phase 4 Demo - Charts",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Column Chart",
                    fill: new SolidFill("4472C4"),
                }),
                new ChartFrame({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 350,
                    type: "column",
                    title: "Quarterly Sales",
                    categories: ["Q1", "Q2", "Q3", "Q4"],
                    series: [
                        { name: "Product A", values: [100, 200, 300, 400] },
                        { name: "Product B", values: [150, 180, 250, 350] },
                    ],
                    showLegend: true,
                    style: 2,
                }),
            ],
        },
        {
            background: new Background({ fill: new SolidFill("F5F5F5") }),
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Pie Chart",
                    fill: new SolidFill("ED7D31"),
                }),
                new ChartFrame({
                    x: 100,
                    y: 120,
                    width: 500,
                    height: 350,
                    type: "pie",
                    title: "Market Share",
                    categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
                    series: [{ name: "Browser", values: [65, 18, 3, 5, 9] }],
                }),
            ],
        },
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Line Chart",
                    fill: new SolidFill("70AD47"),
                }),
                new ChartFrame({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 350,
                    type: "line",
                    title: "Monthly Revenue",
                    categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
                    series: [
                        { name: "2024", values: [10, 15, 13, 20, 25, 30] },
                        { name: "2025", values: [12, 18, 16, 22, 28, 35] },
                    ],
                    style: 2,
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
