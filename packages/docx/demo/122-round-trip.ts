import * as fs from "fs";

import {
  Document,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Packer,
  parseDocx,
  HeadingLevel,
} from "@office-open/docx";
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
  // If text is in a child text node, update that
  const textNode = el.elements?.find((e) => e.type === "text");
  if (textNode) {
    textNode.text = text;
  } else {
    el.text = text;
  }
}

async function main() {
  // 1. Create a document
  const doc = new Document({
    title: "Round-trip Test",
    creator: "Parser Demo",
    sections: [
      {
        children: [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: "Heading 1", bold: true })],
          }),
          new Paragraph({
            children: [new TextRun({ text: "Hello " }), new TextRun({ text: "World", bold: true })],
          }),
          new Paragraph({ text: "Simple paragraph." }),
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ text: "A1" })] }),
                  new TableCell({ children: [new Paragraph({ text: "B1" })] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ text: "A2" })] }),
                  new TableCell({ children: [new Paragraph({ text: "B2" })] }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  // 2. Export to buffer
  const buffer = await Packer.toBuffer(doc);
  console.log(`Generated DOCX: ${buffer.length} bytes`);

  // 3. Parse it back — returns ParsedDocument + w:body Element tree
  const { doc: parsed, body, styles, numbering, partRefs } = parseDocx(new Uint8Array(buffer));

  // 4. Verify Element tree structure
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

  console.log("\n--- Element Tree Verification ---");

  // w:body should have children
  assert("body has children", (body.elements?.length ?? 0) > 0);

  // Count w:p elements
  const paragraphs = body.elements?.filter((e) => e.name === "w:p") ?? [];
  assert("body has 3 paragraphs", paragraphs.length === 3);

  // Count w:tbl elements
  const tables = body.elements?.filter((e) => e.name === "w:tbl") ?? [];
  assert("body has 1 table", tables.length === 1);

  // Verify first paragraph has w:pStyle = Heading1
  const p1 = paragraphs[0];
  const p1Pr = findChild(p1, "w:pPr");
  const p1Style = findChild(p1Pr, "w:pStyle");
  assert("first paragraph has Heading1 style", p1Style?.attributes?.["w:val"] === "Heading1");

  // Verify heading text content
  const headingRun = findChild(p1, "w:r");
  const headingT = findChild(headingRun, "w:t");
  assert("heading text is 'Heading 1'", getText(headingT) === "Heading 1");

  // Verify table has 2 rows
  const tbl = tables[0];
  const tblRows = tbl.elements?.filter((e) => e.name === "w:tr") ?? [];
  assert("table has 2 rows", tblRows.length === 2);

  // 5. Modify Element tree — change heading text
  setText(headingT, "Modified Heading");
  assert("modified heading text", getText(headingT) === "Modified Heading");

  // Write modified document back (body is a child of document Element)
  const documentEl = parsed.get("word/document.xml");
  if (documentEl) parsed.set("word/document.xml", documentEl);

  // 6. Save modified document
  const modifiedBuffer = parsed.save();
  assert("modified buffer non-empty", modifiedBuffer.length > 0);
  console.log(`Modified DOCX: ${modifiedBuffer.length} bytes`);

  // 7. Re-parse modified document and verify
  const { body: modifiedBody } = parseDocx(new Uint8Array(modifiedBuffer));
  const modifiedP1 = modifiedBody.elements?.filter((e) => e.name === "w:p")[0];
  const modifiedRun = findChild(modifiedP1, "w:r");
  const modifiedT = findChild(modifiedRun, "w:t");
  assert("re-parsed modified heading", getText(modifiedT) === "Modified Heading");

  // 8. Test ParsedDocument utility methods
  console.log("\n--- ParsedDocument API ---");
  assert("has word/document.xml", parsed.has("word/document.xml"));
  assert("has word/styles.xml", parsed.has("word/styles.xml"));
  assert("does not have nonexistent", !parsed.has("nonexistent.xml"));

  const allKeys = parsed.keys("word/");
  assert("keys with prefix returns results", allKeys.length > 0);
  console.log(`  word/ keys: ${allKeys.length} files`);

  const rawBinary = parsed.getRaw("word/styles.xml");
  assert("getRaw returns Uint8Array", rawBinary instanceof Uint8Array);

  // 9. Test enhanced parsing — styles, numbering, partRefs
  console.log("\n--- Enhanced Parsing ---");
  assert("styles parsed", !!styles);
  assert("numbering parsed", !!numbering);

  assert("styles root is w:styles", styles!.name === "w:styles");

  console.log(`  headers: ${partRefs.headers.size}, footers: ${partRefs.footers.size}`);
  console.log(`  footnotes: ${partRefs.footnotes ?? "none"}`);
  console.log(`  endnotes: ${partRefs.endnotes ?? "none"}`);
  console.log(`  comments: ${partRefs.comments ?? "none"}`);

  // 9. Unmodified round-trip — parse and save without changes
  const { doc: unmodified } = parseDocx(new Uint8Array(buffer));
  const unmodifiedBuffer = unmodified.save();
  assert("unmodified round-trip produces buffer", unmodifiedBuffer.length > 0);

  // Summary
  console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
  if (fail > 0) process.exit(1);

  // Save files
  fs.writeFileSync("My Document.docx", buffer);
  fs.writeFileSync("My Document (round-trip).docx", modifiedBuffer);
  console.log("\nSaved original and round-trip DOCX files");
}

main().catch(console.error);
