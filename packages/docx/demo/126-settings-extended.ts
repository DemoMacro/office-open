// Extended document settings: mirrorMargins, footnotePr, endnotePr, rsids,
// decimalSymbol, listSeparator, defaultTableStyle, mathPr, readModeInkLockDown,
// bookFoldPrinting, captions, and many onOff settings.

import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
  settings: {
    // Batch A: Boolean onOff settings
    mirrorMargins: true,
    alignBordersAndEdges: true,
    removePersonalInformation: true,
    linkStyles: true,
    autoFormatOverride: false,
    strictFirstAndLastChars: true,
    doNotAutoCompressPictures: true,
    bookFoldPrinting: false,
    bookFoldRevPrinting: false,
    formsDesign: false,
    showEnvelope: false,

    // Batch B: String/number settings
    decimalSymbol: ".",
    listSeparator: ",",
    defaultTableStyle: "TableGrid",
    bookFoldPrintingSheets: 8,
    drawingGridHorizontalSpacing: 120,
    drawingGridVerticalSpacing: 120,
    displayHorizontalDrawingGridEvery: 2,
    displayVerticalDrawingGridEvery: 2,

    // Batch C: Complex substructures
    footnotePr: {
      pos: "pageBottom",
      numFmt: "lowerRoman",
      numStart: 1,
      numRestart: "eachPage",
    },
    endnotePr: {
      pos: "sectEnd",
      numFmt: "lowerLetter",
      numStart: 1,
      numRestart: "eachSect",
    },
    rsids: {
      rsidRoot: "00112233",
      rsids: ["00112233", "44556677", "8899AABB"],
    },
    readModeInkLockDown: {
      actualPg: true,
      w: 12240,
      h: 15840,
      fontSz: 100,
    },
    captions: {
      captions: [
        { name: "Table", pos: "above", numFmt: "upperRoman" },
        { name: "Figure", pos: "below", sep: "hyphen", heading: 1, chapNum: true },
      ],
    },
    mathPr: {
      mathFont: "Cambria Math",
      brkBin: "before",
      smallFrac: true,
      lMargin: 0,
      rMargin: 0,
      defJc: "centerGroup",
      intLim: "subSup",
      naryLim: "subSup",
    },

    // Batch D: Revision tracking
    trackRevisions: true,
    doNotTrackFormatting: true,
    doNotTrackMoves: false,
    revisionView: {
      markup: true,
      comments: true,
      insDel: true,
      formatting: true,
    },
  },

  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "Extended Settings Demo",
              size: 32,
            }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun(
              "This document demonstrates extended settings including footnotePr, endnotePr, rsids, mathPr, captions, and more.",
            ),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
