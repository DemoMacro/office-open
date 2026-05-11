import * as fs from "fs";

import { Presentation, Shape, Packer, Paragraph, TextRun, parsePptx } from "@office-open/pptx";
import type { Element } from "@office-open/xml";
import { findChild } from "@office-open/xml";

function getText(el: Element | undefined): string {
    if (!el) return "";
    if (el.text != null) return String(el.text);
    return (
        el.elements
            ?.filter((e) => e.type === "text")
            .map((e) => String(e.text ?? ""))
            .join("") ?? ""
    );
}

function setText(el: Element | undefined, text: string): void {
    if (!el) return;
    const textNode = el.elements?.find((e) => e.type === "text");
    if (textNode) {
        textNode.text = text;
    } else {
        el.text = text;
    }
}

async function main() {
    // 1. Create a presentation
    const pres = new Presentation({
        title: "Round-trip Test",
        creator: "Parser Demo",
        slides: [
            {
                children: [
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
                    new Shape({
                        x: 100,
                        y: 400,
                        width: 500,
                        height: 200,
                        paragraphs: [
                            new Paragraph({
                                children: [
                                    new TextRun({ text: "Bold", bold: true }),
                                    new TextRun({ text: ", Italic", italic: true }),
                                    new TextRun({ text: ", Underline", underline: "SINGLE" }),
                                ],
                            }),
                            new Paragraph({
                                children: [new TextRun({ text: "Second paragraph" })],
                            }),
                        ],
                    }),
                ],
            },
            {
                children: [
                    new Shape({
                        x: 200,
                        y: 200,
                        width: 300,
                        height: 150,
                        text: "Slide 2",
                        geometry: "diamond",
                        fill: "ED7D31",
                    }),
                ],
            },
        ],
    });

    // 2. Export to buffer
    const buffer = await Packer.toBuffer(pres);
    console.log(`Generated PPTX: ${buffer.length} bytes`);

    // 3. Parse it back — returns ParsedDocument + slide paths
    const {
        doc: parsed,
        slidePaths,
        slideMasterPaths,
        slideLayoutPaths,
        notesSlidePaths,
        slideWidth,
        slideHeight,
    } = parsePptx(new Uint8Array(buffer));

    // 4. Verify parsed data
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

    console.log("\n--- ParsedDocument Verification ---");

    // Slide paths
    assert("2 slides", slidePaths.length === 2);
    assert("slide 1 path", slidePaths[0] === "ppt/slides/slide1.xml");
    assert("slide 2 path", slidePaths[1] === "ppt/slides/slide2.xml");

    // Slide dimensions
    assert("has slideWidth", slideWidth !== undefined);
    assert("has slideHeight", slideHeight !== undefined);
    if (slideWidth && slideHeight) {
        console.log(`  slide size: ${slideWidth} x ${slideHeight}`);
    }

    // 5. Inspect slide 1 Element tree
    console.log("\n--- Slide 1 Element Tree ---");
    const slide1 = parsed.get(slidePaths[0]);
    assert("slide 1 parsed", !!slide1);

    if (slide1) {
        // Shapes are nested under p:cSld/p:spTree
        const cSld = findChild(slide1, "p:cSld");
        const spTree = cSld ? findChild(cSld, "p:spTree") : null;
        const shapes = spTree?.elements?.filter((e) => e.name === "p:sp") ?? [];
        assert("slide 1 has 3 shapes", shapes.length === 3);

        // Verify first shape has text "Hello PPTX"
        const sp1TxBody = findChild(shapes[0], "p:txBody");
        const sp1FirstPara = sp1TxBody?.elements?.find((e) => e.name === "a:p");
        const sp1Run = sp1FirstPara?.elements?.find((e) => e.name === "a:r");
        const sp1T = findChild(sp1Run, "a:t");
        assert("shape 1 text is 'Hello PPTX'", getText(sp1T) === "Hello PPTX");

        // Verify third shape has 2 paragraphs (rich text)
        const sp3TxBody = findChild(shapes[2], "p:txBody");
        const sp3Paras = sp3TxBody?.elements?.filter((e) => e.name === "a:p") ?? [];
        assert("shape 3 has 2 paragraphs", sp3Paras.length === 2);

        // Modify first shape text
        setText(sp1T, "Modified PPTX");

        // Write modified slide back
        parsed.set(slidePaths[0], slide1);
    }

    // 6. Save modified document
    const modifiedBuffer = parsed.save();
    assert("modified buffer non-empty", modifiedBuffer.length > 0);
    console.log(`Modified PPTX: ${modifiedBuffer.length} bytes`);

    // 7. Re-parse and verify modification
    const { doc: reparsed, slidePaths: reparsedPaths } = parsePptx(new Uint8Array(modifiedBuffer));
    const reparsedSlide1 = reparsed.get(reparsedPaths[0]);
    const reparsedCSld = findChild(reparsedSlide1!, "p:cSld");
    const reparsedSpTree = reparsedCSld ? findChild(reparsedCSld, "p:spTree") : null;
    const reparsedSp1 = reparsedSpTree?.elements?.find((e) => e.name === "p:sp");
    const reparsedTxBody = findChild(reparsedSp1!, "p:txBody");
    const reparsedFirstPara = reparsedTxBody?.elements?.find((e) => e.name === "a:p");
    const reparsedRun = reparsedFirstPara?.elements?.find((e) => e.name === "a:r");
    const reparsedT = findChild(reparsedRun, "a:t");
    assert("re-parsed modified text", getText(reparsedT) === "Modified PPTX");

    // 8. Test ParsedDocument utility methods
    console.log("\n--- ParsedDocument API ---");
    assert("has ppt/presentation.xml", parsed.has("ppt/presentation.xml"));
    assert("has ppt/theme/theme1.xml", parsed.has("ppt/theme/theme1.xml"));

    const pptKeys = parsed.keys("ppt/slides/");
    assert("keys returns slide paths", pptKeys.length >= 2);
    console.log(`  ppt/slides/ keys: ${pptKeys.join(", ")}`);

    const presEl = parsed.get("ppt/presentation.xml");
    assert("get presentation.xml returns Element", !!presEl);
    assert("presentation has sldSz", !!findChild(presEl!, "p:sldSz"));

    // 9. Test enhanced parsing
    console.log("\n--- Enhanced Parsing ---");
    assert("has slide masters", slideMasterPaths.length >= 1);
    assert("has slide layouts", slideLayoutPaths.length >= 1);
    console.log(`  slideMasters: ${slideMasterPaths.join(", ")}`);
    console.log(`  slideLayouts: ${slideLayoutPaths.join(", ")}`);
    console.log(`  notesSlides: ${notesSlidePaths.length}`);

    // Summary
    console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
    if (fail > 0) process.exit(1);

    // Save files
    fs.writeFileSync("My Presentation.pptx", buffer);
    fs.writeFileSync("My Presentation (round-trip).pptx", modifiedBuffer);
    console.log("\nSaved original and round-trip PPTX files");
}

main().catch(console.error);
