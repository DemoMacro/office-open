import * as fs from "fs";

import { Presentation, Shape, Packer, Paragraph, Run, parsePptx } from "@office-open/pptx";
import type { ShapeJson } from "@office-open/pptx";

function asShape(v: unknown): ShapeJson {
    return v as ShapeJson;
}

async function main() {
    // 1. Create a presentation with diverse content
    const pres = new Presentation({
        title: "Round-trip Test",
        creator: "Parser Demo",
        slides: [
            {
                children: [
                    // Simple shape with text
                    new Shape({
                        x: 100,
                        y: 100,
                        width: 400,
                        height: 200,
                        text: "Hello PPTX",
                        geometry: "rect",
                        fill: "4472C4",
                    }),
                    // Shape with rich text paragraphs
                    new Shape({
                        x: 200,
                        y: 400,
                        width: 500,
                        height: 100,
                        paragraphs: [
                            new Paragraph({
                                children: [
                                    new Run({ text: "Rich text: ", bold: true }),
                                    new Run({ text: "italic", italic: true, fontSize: 18 }),
                                    new Run({ text: ", colored", fill: "FF0000" }),
                                ],
                            }),
                        ],
                    }),
                    // Rounded rectangle
                    new Shape({
                        x: 600,
                        y: 100,
                        width: 200,
                        height: 150,
                        text: "Rounded",
                        geometry: "roundRect",
                        fill: "70AD47",
                    }),
                ],
            },
        ],
    });

    // 2. Export to buffer
    const buffer = await Packer.toBuffer(pres);
    console.log(`Generated PPTX: ${buffer.length} bytes`);

    // 3. Parse it back
    const json = await parsePptx(new Uint8Array(buffer));

    // 4. Verify round-trip
    let pass = 0;
    let fail = 0;
    const assert = (label: string, condition: boolean) => {
        if (condition) {
            pass++;
            console.log(`  PASS: ${label}`);
        } else {
            fail++;
            console.log(`  FAIL: ${label}`);
        }
    };

    console.log("\n--- Core Properties ---");
    assert("title", json.title === "Round-trip Test");
    assert("creator", json.creator === "Parser Demo");

    console.log("\n--- Presentation ---");
    assert("slide width 9144000 (10 inches)", json.slideWidth === 9144000);
    assert("slide height 6858000 (7.5 inches)", json.slideHeight === 6858000);

    console.log("\n--- Slides ---");
    assert("1 slide", json.slides.length === 1);

    const slide = json.slides[0];
    assert("slide has 3 children", slide.children.length === 3);

    // Shape 1: Simple rect with text
    const s1 = asShape(slide.children[0]);
    assert("shape 1 type", s1.$type === "shape");
    assert("shape 1 geometry rect", s1.geometry === "rect");
    assert("shape 1 has id", typeof s1.id === "number");
    assert("shape 1 has name", typeof s1.name === "string" && s1.name!.startsWith("Shape"));
    assert("shape 1 x", s1.x === 952500);
    assert("shape 1 y", s1.y === 952500);
    assert("shape 1 width", s1.width === 3810000);
    assert("shape 1 height", s1.height === 1905000);
    const s1Fill = s1.fill as Record<string, unknown>;
    assert("shape 1 fill type solid", s1Fill?.type === "solid");
    assert("shape 1 fill color", s1Fill?.color === "4472C4");
    const s1Paras = s1.paragraphs;
    assert("shape 1 has 1 paragraph", s1Paras?.length === 1);
    assert("shape 1 text", s1Paras?.[0]?.text === "Hello PPTX");

    // Shape 2: Rich text
    const s2 = asShape(slide.children[1]);
    assert("shape 2 type", s2.$type === "shape");
    assert("shape 2 width", s2.width === 4762500);
    const s2Paras = s2.paragraphs;
    assert("shape 2 has 1 paragraph", s2Paras?.length === 1);
    const s2Runs = s2Paras?.[0]?.children;
    assert("shape 2 has 3 runs", s2Runs?.length === 3);
    assert("run 1 text", s2Runs?.[0]?.text === "Rich text: ");
    assert("run 1 bold", s2Runs?.[0]?.bold === true);
    assert("run 2 text", s2Runs?.[1]?.text === "italic");
    assert("run 2 italic", s2Runs?.[1]?.italic === true);
    assert("run 2 fontSize", s2Runs?.[1]?.fontSize === 1800);
    assert("run 3 text", s2Runs?.[2]?.text === ", colored");
    const run3Fill = s2Runs?.[2]?.fill as Record<string, unknown>;
    assert("run 3 fill type solid", run3Fill?.type === "solid");
    assert("run 3 fill color", run3Fill?.color === "FF0000");

    // Shape 3: Rounded rectangle
    const s3 = asShape(slide.children[2]);
    assert("shape 3 type", s3.$type === "shape");
    assert("shape 3 geometry roundRect", s3.geometry === "roundRect");
    const s3Fill = s3.fill as Record<string, unknown>;
    assert("shape 3 fill color", s3Fill?.color === "70AD47");
    assert("shape 3 text", s3.paragraphs?.[0]?.text === "Rounded");

    // Summary
    console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
    if (fail > 0) process.exit(1);

    // Save to file
    fs.writeFileSync("My Presentation.pptx", buffer);
    console.log("\nSaved to My Presentation.pptx");
}

main().catch(console.error);
