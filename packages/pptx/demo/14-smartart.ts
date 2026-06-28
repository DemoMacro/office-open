import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "SmartArt Demo",
  creator: "Demo",
  slides: [
    // --- 1. List Layouts ---
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "List Layouts", size: 32, bold: true }],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "1.3cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "7.9cm",
            nodes: [
              { text: "Item A", children: [{ text: "Sub A1" }, { text: "Sub A2" }] },
              { text: "Item B", children: [{ text: "Sub B1" }] },
              { text: "Item C" },
            ],
            layout: "default",
            style: "simple1",
            color: "accent1_2",
          },
        },
        {
          smartart: {
            x: "12.7cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "7.9cm",
            nodes: [{ text: "Step 1" }, { text: "Step 2" }, { text: "Step 3" }],
            layout: "list1",
            style: "simple2",
            color: "accent2_2",
          },
        },
      ],
    },

    // --- 2. Process Layouts ---
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Process Layouts", size: 32, bold: true }],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "1.3cm",
            y: "3.2cm",
            width: "22.5cm",
            height: "7.9cm",
            nodes: [
              { text: "Plan" },
              { text: "Design" },
              { text: "Build" },
              { text: "Test" },
              { text: "Deploy" },
            ],
            layout: "process1",
            style: "moderate1",
            color: "accent3_2",
          },
        },
      ],
    },
    {
      children: [
        {
          smartart: {
            x: "1.3cm",
            y: "1.3cm",
            width: "22.5cm",
            height: "9.3cm",
            nodes: [
              { text: "Phase 1" },
              { text: "Phase 2" },
              { text: "Phase 3" },
              { text: "Phase 4" },
            ],
            layout: "chevron1",
            style: "professional1",
            color: "accent4_2",
          },
        },
      ],
    },
    {
      children: [
        {
          smartart: {
            x: "1.3cm",
            y: "1.3cm",
            width: "22.5cm",
            height: "9.3cm",
            nodes: [
              { text: "Start" },
              { text: "Process" },
              { text: "Review" },
              { text: "Complete" },
            ],
            layout: "arrow1",
            style: "polished1",
            color: "accent5_2",
          },
        },
      ],
    },

    // --- 3. Cycle Layouts ---
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Cycle Layouts", size: 32, bold: true }],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "1.3cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [
              { text: "Research" },
              { text: "Design" },
              { text: "Develop" },
              { text: "Test" },
            ],
            layout: "cycle1",
            style: "cartoon1",
            color: "colorful1",
          },
        },
        {
          smartart: {
            x: "12.7cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [
              { text: "Plan" },
              { text: "Do" },
              { text: "Check" },
              { text: "Act" },
              { text: "Improve" },
            ],
            layout: "cycle2",
            style: "powdery1",
            color: "colorful2",
          },
        },
      ],
    },

    // --- 4. Hierarchy Layouts ---
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Hierarchy Layouts",
                      size: 32,
                      bold: true,
                    },
                  ],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "10.6cm",
            nodes: [
              {
                text: "CEO",
                children: [
                  {
                    text: "VP Engineering",
                    children: [{ text: "Frontend" }, { text: "Backend" }],
                  },
                  {
                    text: "VP Marketing",
                    children: [{ text: "Brand" }, { text: "Growth" }],
                  },
                ],
              },
            ],
            layout: "hierarchy1",
            style: "moderate2",
            color: "accent1_2",
          },
        },
      ],
    },
    {
      children: [
        {
          smartart: {
            x: "1.3cm",
            y: "1.3cm",
            width: "15.9cm",
            height: "10.6cm",
            nodes: [
              {
                text: "Root",
                children: [
                  {
                    text: "Branch A",
                    children: [{ text: "Leaf A1" }, { text: "Leaf A2" }],
                  },
                  { text: "Branch B", children: [{ text: "Leaf B1" }] },
                  { text: "Branch C" },
                ],
              },
            ],
            layout: "hierarchy2",
            style: "polished2",
            color: "colorful3",
          },
        },
      ],
    },

    // --- 5. Relationship Layouts ---
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Relationship / Other Layouts",
                      size: 32,
                      bold: true,
                    },
                  ],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "1.3cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [{ text: "Set A" }, { text: "Set B" }, { text: "Set C" }],
            layout: "venn1",
            style: "burnt1",
            color: "accent6_2",
          },
        },
        {
          smartart: {
            x: "12.7cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [
              { text: "Leads" },
              { text: "Prospects" },
              { text: "Qualified" },
              { text: "Negotiation" },
              { text: "Closed" },
            ],
            layout: "funnel1",
            style: "professional2",
            color: "dark1",
          },
        },
      ],
    },

    // --- 6. Pyramid Layouts ---
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Pyramid Layouts", size: 32, bold: true }],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "1.3cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [
              { text: "Vision" },
              { text: "Strategy" },
              { text: "Tactics" },
              { text: "Operations" },
            ],
            layout: "pyramid1",
            style: "moderate3",
            color: "accent2_2",
          },
        },
        {
          smartart: {
            x: "12.7cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [{ text: "All" }, { text: "Most" }, { text: "Some" }, { text: "Few" }],
            layout: "pyramid2",
            style: "cartoon2",
            color: "colorful4",
          },
        },
      ],
    },

    // --- 7. Matrix / Radial / Other Layouts ---
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Matrix / Radial / Other Layouts",
                      size: 32,
                      bold: true,
                    },
                  ],
                },
              ],
            },
          },
        },
        {
          smartart: {
            x: "1.3cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [
              { text: "Q1: Plan" },
              { text: "Q2: Execute" },
              { text: "Q3: Review" },
              { text: "Q4: Deliver" },
            ],
            layout: "matrix1",
            style: "professional3",
            color: "primary1",
          },
        },
        {
          smartart: {
            x: "12.7cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [{ text: "Core" }, { text: "Team" }, { text: "Tools" }, { text: "Process" }],
            layout: "radial1",
            style: "burnt2",
            color: "gray1",
          },
        },
      ],
    },
    {
      children: [
        {
          smartart: {
            x: "1.3cm",
            y: "1.3cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [{ text: "Input" }, { text: "Process" }, { text: "Output" }],
            layout: "balance1",
            style: "powdery2",
            color: "primary2",
          },
        },
        {
          smartart: {
            x: "12.7cm",
            y: "1.3cm",
            width: "10.6cm",
            height: "9.3cm",
            nodes: [{ text: "Part A" }, { text: "Part B" }, { text: "Part C" }, { text: "Part D" }],
            layout: "gear1",
            style: "polished3",
            color: "accent3_2",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
