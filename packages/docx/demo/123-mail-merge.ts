// Mail merge: configure a document as a merge template linked to a data source.
// Word will show "This document is a merge template" when opened.
// Use Mailings > Start Mail Merge to preview or complete the merge.

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
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
        {
          paragraph: {
            children: ["Mail Merge Template"],
          },
        },
        {
          paragraph: {
            children: ["This document is configured as a mail merge template linked to data.xlsx."],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
