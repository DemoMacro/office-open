// Document settings features: view, zoom, write protection, display background shape,
// font embedding, document variables, mail merge

import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
  background: {
    color: "C45911",
  },
  displayBackgroundShape: true,
  view: "print",
  zoom: {
    percent: 150,
  },
  embedTrueTypeFonts: true,
  saveSubsetFonts: true,
  docVars: [
    { name: "Title", val: "Settings Demo" },
    { name: "Version", val: "1.0" },
    { name: "Author", val: "Test User" },
  ],
  mailMerge: {
    mainDocumentType: "formLetters",
    dataType: "spreadsheet",
    connectString: "DSN=Excel Files;DBQ=data.xlsx",
    query: "SELECT * FROM `Sheet1$`",
    dataSource: "data.xlsx",
    destination: "newDocument",
    addressFieldName: "Email",
    mailSubject: "Monthly Report",
    linkToQuery: true,
    doNotSuppressBlankLines: true,
    odso: {
      udl: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=data.xlsx",
      table: "Sheet1$",
      type: "database",
      fieldMapData: [
        { type: "dbColumn", name: "FirstName", mappedName: "First Name", column: 0 },
        { type: "dbColumn", name: "LastName", mappedName: "Last Name", column: 1 },
      ],
    },
  },
  sections: [
    {
      children: [
        new Paragraph({
          children: [new TextRun("Document Settings Demo")],
        }),
        new Paragraph({
          children: [new TextRun("This document opens in Print Layout view at 150% zoom.")],
        }),
        new Paragraph({
          children: [
            new TextRun(
              "The background color is displayed in print layout because displayBackgroundShape is enabled.",
            ),
          ],
        }),
        new Paragraph({
          children: [new TextRun("TrueType fonts are embedded, and only used subsets are saved.")],
        }),
        new Paragraph({
          children: [
            new TextRun("Document variables (Title, Version, Author) are stored in settings.xml."),
          ],
        }),
        new Paragraph({
          children: [new TextRun("Mail merge is configured to use a spreadsheet data source.")],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
