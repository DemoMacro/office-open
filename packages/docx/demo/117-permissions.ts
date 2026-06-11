// Demo: Permission Ranges (Document Protection - Editable Regions)
// Demonstrates PermStart and PermEnd for defining editable regions.
// Note: documentProtection must be enabled in features for permissions to take effect.

import { writeFileSync } from "node:fs";

import { EditGroupType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  features: {
    trackRevisions: true,
    documentProtection: {
      edit: "readOnly",
      password: "123",
    },
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Permission Ranges Demo", bold: true, size: 32 }],
            spacing: { after: 400 },
          },
        },
        {
          paragraph: {
            children: ["Password: 123 | Protection: Read-only (with editable ranges)"],
          },
        },

        { paragraph: "" },

        {
          paragraph: {
            children: [{ bold: true, text: "1. Everyone Editable Range", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: ["This text is read-only (protected)."],
          },
        },
        {
          paragraph: {
            children: [
              { permStart: { id: 0, editGroup: EditGroupType.EVERYONE } },
              "Everyone can edit this text.",
              { permEnd: 0 },
            ],
          },
        },
        {
          paragraph: {
            children: ["This text is read-only again."],
          },
        },

        { paragraph: "" },

        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "2. Single User Editable Range",
                size: 28,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              { permStart: { id: 1, ed: "john@example.com" } },
              "Only john@example.com can edit this text.",
              { permEnd: 1 },
            ],
          },
        },

        { paragraph: "" },

        {
          paragraph: {
            children: [{ bold: true, text: "3. Editors Group Range", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              { permStart: { id: 2, editGroup: EditGroupType.EDITORS } },
              "Only editors can modify this content.",
              { permEnd: 2 },
            ],
          },
        },

        { paragraph: "" },

        {
          paragraph: {
            children: [{ bold: true, text: "4. Current User Only Range", size: 28 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              { permStart: { id: 3, editGroup: EditGroupType.CURRENT } },
              "Only the current user can edit this.",
              { permEnd: 3 },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
