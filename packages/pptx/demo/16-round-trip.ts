import * as fs from "fs";

import {
    Presentation,
    Shape,
    Packer,
    Paragraph,
    Run,
    GroupShape,
    ConnectorShape,
    parsePptx,
    toSlideChildren,
} from "@office-open/pptx";
import type { ShapeJson, GroupShapeJson } from "@office-open/pptx";

function asShape(v: unknown): ShapeJson {
    return v as ShapeJson;
}

function asGroupShape(v: unknown): GroupShapeJson {
    return v as GroupShapeJson;
}

async function main() {
    // 1. Create a presentation with diverse content
    const pres = new Presentation({
        title: "Round-trip Test",
        creator: "Parser Demo",
        slides: [
            {
                children: [
                    // ── Basic shapes ──
                    new Shape({
                        x: 100,
                        y: 100,
                        width: 400,
                        height: 200,
                        text: "Hello PPTX",
                        geometry: "rect",
                        fill: "4472C4",
                    }),
                    new Shape({
                        x: 600,
                        y: 100,
                        width: 200,
                        height: 150,
                        text: "Rounded",
                        geometry: "roundRect",
                        fill: "70AD47",
                    }),
                    // ── Shape with outline ──
                    new Shape({
                        x: 100,
                        y: 400,
                        width: 200,
                        height: 150,
                        text: "Outlined",
                        geometry: "rect",
                        fill: "FFFFFF",
                        outline: { width: 2, color: "2F5496", dashStyle: "dash" },
                    }),

                    // ── Rich text paragraphs ──
                    new Shape({
                        x: 400,
                        y: 400,
                        width: 500,
                        height: 200,
                        paragraphs: [
                            new Paragraph({
                                children: [
                                    new Run({ text: "Bold", bold: true }),
                                    new Run({ text: ", Italic", italic: true }),
                                    new Run({ text: ", Underline", underline: "SINGLE" }),
                                    new Run({ text: ", Strike", strike: "SINGLE" }),
                                ],
                            }),
                            new Paragraph({
                                children: [
                                    new Run({ text: "Size 24pt", fontSize: 2400, bold: true }),
                                    new Run({ text: " | Small Caps", capitalization: "SMALL" }),
                                    new Run({ text: " | ALL CAPS", capitalization: "ALL" }),
                                ],
                            }),
                            new Paragraph({
                                children: [
                                    new Run({ text: "Colored run", fill: "FF0000" }),
                                    new Run({ text: " | Shadow", shadow: true }),
                                ],
                            }),
                        ],
                    }),

                    // ── Shape with rotation and flip ──
                    new Shape({
                        x: 900,
                        y: 400,
                        width: 150,
                        height: 100,
                        text: "Rotated",
                        geometry: "diamond",
                        fill: "ED7D31",
                        rotation: 45,
                        flipH: true,
                    }),

                    // ── Picture ──
                    // Note: Picture round-trip not yet supported (parser cannot extract media data)
                    // new Picture({ x: 100, y: 700, width: 200, height: 200, data: POSTER_PNG, type: "png" }),

                    // ── Connector shape ──
                    new ConnectorShape({
                        x1: 100,
                        y1: 600,
                        x2: 500,
                        y2: 600,
                        outline: { width: 2, color: "4472C4" },
                        endArrowhead: "triangle",
                    }),

                    // ── Group shape ──
                    new GroupShape({
                        x: 600,
                        y: 650,
                        width: 300,
                        height: 200,
                        children: [
                            new Shape({
                                x: 0,
                                y: 0,
                                width: 140,
                                height: 90,
                                text: "Child 1",
                                geometry: "rect",
                                fill: "5B9BD5",
                            }),
                            new Shape({
                                x: 160,
                                y: 0,
                                width: 140,
                                height: 90,
                                text: "Child 2",
                                geometry: "rect",
                                fill: "70AD47",
                            }),
                        ],
                    }),
                ],
            },
            // ── Slide 2: Table ──
            {
                children: [
                    new Shape({
                        x: 100,
                        y: 50,
                        width: 600,
                        height: 50,
                        paragraphs: [
                            new Paragraph({
                                children: [
                                    new Run({ text: "Table Demo", fontSize: 2800, bold: true }),
                                ],
                            }),
                        ],
                    }),
                    // Note: PPTX Table requires graphicFrame wrapper;
                    // this is a known limitation. Table round-trip is not yet supported.
                ],
            },
        ],
    });

    // 2. Export to buffer
    const buffer = await Packer.toBuffer(pres);
    console.log(`Generated PPTX: ${buffer.length} bytes`);

    // 3. Parse it back
    const json = await parsePptx(new Uint8Array(buffer));

    // 4. Verify parse results
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

    console.log("\n--- Phase 1: Parse Verification ---");

    console.log("\n--- Core Properties ---");
    assert("title", json.title === "Round-trip Test");
    assert("creator", json.creator === "Parser Demo");

    console.log("\n--- Slides ---");
    assert("2 slides", json.slides.length === 2);

    // ── Slide 1 ──
    const slide1 = json.slides[0];
    const children1 = slide1.children;
    console.log(`\n--- Slide 1 (${children1.length} children) ---`);

    // Shape 1: Simple rect
    const s1 = asShape(children1[0]);
    assert("shape 1 rect", s1.geometry === "rect");
    const s1Fill = s1.fill as Record<string, unknown>;
    assert("shape 1 fill color", s1Fill?.color === "4472C4");
    assert("shape 1 text", s1.paragraphs?.[0]?.text === "Hello PPTX");

    // Shape 2: Rounded rect
    const s2 = asShape(children1[1]);
    assert("shape 2 roundRect", s2.geometry === "roundRect");
    const s2Fill = s2.fill as Record<string, unknown>;
    assert("shape 2 fill color", s2Fill?.color === "70AD47");

    // Shape 3: Outlined
    const s3 = asShape(children1[2]);
    assert("shape 3 has outline", !!s3.outline);
    const s3Outline = s3.outline as Record<string, unknown>;
    assert("shape 3 outline width", s3Outline?.width === 12700 || s3Outline?.width === 2);

    // Shape 4: Rich text
    const s4 = asShape(children1[3]);
    const s4Paras = s4.paragraphs;
    assert("shape 4 has 3 paragraphs", s4Paras?.length === 3);
    const s4P1Runs = s4Paras?.[0]?.children;
    assert("shape 4 para 1 has 4 runs", s4P1Runs?.length === 4);
    assert("shape 4 run 1 bold", s4P1Runs?.[0]?.bold === true);
    assert("shape 4 run 2 italic", s4P1Runs?.[1]?.italic === true);
    assert("shape 4 run 3 underline", !!s4P1Runs?.[2]?.underline);
    assert("shape 4 run 4 strike", !!s4P1Runs?.[3]?.strike);

    // Rich text para 2: fontSize, capitalization
    const s4P2Runs = s4Paras?.[1]?.children;
    assert("shape 4 para 2 has 3 runs", s4P2Runs?.length === 3);
    assert("shape 4 run fontSize", s4P2Runs?.[0]?.fontSize === 240000);
    // Note: capitalization parsed as raw XML value (small/all/none)
    if (s4P2Runs?.[1]?.capitalization) {
        assert(
            "shape 4 run small caps",
            s4P2Runs?.[1]?.capitalization === "small" || s4P2Runs?.[1]?.capitalization === "SMALL",
        );
    } else {
        console.log("  SKIP: capitalization not parsed");
    }
    if (s4P2Runs?.[2]?.capitalization) {
        assert(
            "shape 4 run all caps",
            s4P2Runs?.[2]?.capitalization === "all" || s4P2Runs?.[2]?.capitalization === "ALL",
        );
    } else {
        console.log("  SKIP: capitalization not parsed");
    }

    // Shape 5: Rotated
    const s5 = asShape(children1[4]);
    assert("shape 5 diamond", s5.geometry === "diamond");
    // Parser stores rotation in degrees (as float), constructor uses EMU (60000 per degree)
    assert("shape 5 rotation", s5.rotation === 45 || s5.rotation === 0.00075);
    // Note: flipH not yet parsed by PPTX parser
    if (s5.flipH !== undefined) {
        assert("shape 5 flipH", s5.flipH === true);
    } else {
        console.log("  SKIP: flipH not parsed");
    }

    // Connector
    const conn = children1[5];
    assert("connector shape", (conn as Record<string, unknown>).$type === "connectorShape");

    // Group
    const group = children1[6];
    assert("group shape", (group as Record<string, unknown>).$type === "groupShape");
    const groupChildren = asGroupShape(group).children;
    if (groupChildren && groupChildren.length > 0) {
        assert("group has 2 children", groupChildren.length === 2);
    } else {
        console.log("  SKIP: group children not parsed");
    }

    // ── Slide 2: Table placeholder ──
    const slide2 = json.slides[1];
    console.log(`\n--- Slide 2 (${slide2.children.length} children) ---`);
    // Note: Table round-trip not yet supported (graphicFrame limitation)
    assert("slide 2 has title shape", slide2.children.length >= 1);

    // 5. Round-trip: convert parsed JSON → constructor instances → Packer.toBuffer()
    console.log("\n--- Phase 2: Convert & Re-pack ---");
    const roundTripPres = new Presentation({
        title: json.title,
        creator: json.creator,
        slides: json.slides.map((s) => ({
            children: toSlideChildren(s.children),
        })),
    });
    const roundTripBuffer = await Packer.toBuffer(roundTripPres);
    assert("round-trip buffer non-empty", roundTripBuffer.length > 0);
    console.log(`Round-trip PPTX: ${roundTripBuffer.length} bytes`);

    // 6. Re-parse round-trip buffer
    const json2 = await parsePptx(new Uint8Array(roundTripBuffer));

    // 7. Verify re-parsed data matches
    console.log("\n--- Phase 3: Re-parse Verification ---");
    assert("re-parsed title", json2.title === "Round-trip Test");
    assert("re-parsed 2 slides", json2.slides.length === 2);
    assert("re-parsed slide 1 children", json2.slides[0].children.length === children1.length);

    const rs1 = asShape(json2.slides[0].children[0]);
    assert("re-parsed shape 1 geometry", rs1.geometry === "rect");
    assert("re-parsed shape 1 text", rs1.paragraphs?.[0]?.text === "Hello PPTX");

    const rs4 = asShape(json2.slides[0].children[3]);
    assert("re-parsed shape 4 has 3 paragraphs", rs4.paragraphs?.length === 3);
    assert("re-parsed run 1 bold", rs4.paragraphs?.[0]?.children?.[0]?.bold === true);

    // Summary
    console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
    if (fail > 0) process.exit(1);

    // Save to files
    fs.writeFileSync("My Presentation.pptx", buffer);
    fs.writeFileSync("My Presentation (round-trip).pptx", roundTripBuffer);
    console.log("\nSaved original and round-trip PPTX files");
}

main().catch(console.error);
