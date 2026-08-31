import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Animation Demo",
  creator: "Demo",
  slides: [
    // Slide 1: Entrance animations
    {
      children: [
        {
          shape: {
            id: 2,
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Entrance Animations" },
            properties: {
              fill: "4472C4",
            },
          },
        },
        {
          shape: {
            id: 3,
            x: "1.3cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Appear" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
        {
          shape: {
            id: 4,
            x: "10.6cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Fly In (from left)" },
            properties: {
              fill: "70AD47",
            },
          },
        },
        {
          shape: {
            id: 5,
            x: "1.3cm",
            y: "6.6cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Wipe (down)" },
            properties: {
              fill: "FFC000",
            },
          },
        },
        {
          shape: {
            id: 6,
            x: "10.6cm",
            y: "6.6cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Dissolve" },
            properties: {
              fill: "5B9BD5",
            },
          },
        },
        {
          shape: {
            id: 7,
            x: "1.3cm",
            y: "10.1cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Zoom In" },
            properties: {
              fill: "BF8F00",
            },
          },
        },
        {
          shape: {
            id: 8,
            x: "10.6cm",
            y: "10.1cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Split (horizontal)" },
            properties: {
              fill: "7030A0",
            },
          },
        },
      ],
      animations: [
        { shapeId: 2, type: "fade", duration: 800 },
        { shapeId: 3, type: "appear" },
        { shapeId: 4, type: "fly", direction: "left", duration: 600 },
        { shapeId: 5, type: "wipe", direction: "down" },
        { shapeId: 6, type: "dissolve", duration: 1000 },
        { shapeId: 7, type: "zoom" },
        { shapeId: 8, type: "split", direction: "horizontal" },
      ],
    },

    // Slide 2: Exit animations
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Exit Animations" },
            properties: {
              fill: "C00000",
            },
          },
        },
        {
          shape: {
            id: 2,
            x: "1.3cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Fade Out" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
        {
          shape: {
            id: 3,
            x: "10.6cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Fly Out (right)" },
            properties: {
              fill: "70AD47",
            },
          },
        },
        {
          shape: {
            id: 4,
            x: "1.3cm",
            y: "6.6cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Wipe Out (up)" },
            properties: {
              fill: "FFC000",
            },
          },
        },
        {
          shape: {
            id: 5,
            x: "10.6cm",
            y: "6.6cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Dissolve Out" },
            properties: {
              fill: "5B9BD5",
            },
          },
        },
        {
          shape: {
            id: 6,
            x: "1.3cm",
            y: "10.1cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Zoom Out" },
            properties: {
              fill: "BF8F00",
            },
          },
        },
        {
          shape: {
            id: 7,
            x: "10.6cm",
            y: "10.1cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Split Out" },
            properties: {
              fill: "7030A0",
            },
          },
        },
      ],
      animations: [
        { shapeId: 2, type: "fade", class: "exit", duration: 800 },
        { shapeId: 3, type: "fly", class: "exit", direction: "right", duration: 600 },
        { shapeId: 4, type: "wipe", class: "exit", direction: "up" },
        { shapeId: 5, type: "dissolve", class: "exit", duration: 1000 },
        { shapeId: 6, type: "zoom", class: "exit" },
        { shapeId: 7, type: "split", class: "exit", direction: "horizontal" },
      ],
    },

    // Slide 3: Emphasis animations
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Emphasis Animations" },
            properties: {
              fill: "548235",
            },
          },
        },
        {
          shape: {
            id: 2,
            x: "1.3cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Grow/Shrink" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
        {
          shape: {
            id: 3,
            x: "10.6cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Spin" },
            properties: {
              fill: "70AD47",
            },
          },
        },
        {
          shape: {
            id: 4,
            x: "1.3cm",
            y: "6.6cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Color Change" },
            properties: {
              fill: "FFC000",
            },
          },
        },
        {
          shape: {
            id: 5,
            x: "10.6cm",
            y: "6.6cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Transparency" },
            properties: {
              fill: "5B9BD5",
            },
          },
        },
        {
          shape: {
            id: 6,
            x: "1.3cm",
            y: "10.1cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Pulse" },
            properties: {
              fill: "BF8F00",
            },
          },
        },
        {
          shape: {
            id: 7,
            x: "10.6cm",
            y: "10.1cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Bold Flash" },
            properties: {
              fill: "7030A0",
            },
          },
        },
      ],
      animations: [
        {
          shapeId: 2,
          class: "emphasis",
          emphasisType: "growShrink",
          zoomContents: true,
          duration: 800,
        },
        { shapeId: 3, class: "emphasis", emphasisType: "spin", duration: 1000 },
        {
          shapeId: 4,
          class: "emphasis",
          emphasisType: "colorChange",
          color: "FF0000",
          duration: 800,
        },
        { shapeId: 5, class: "emphasis", emphasisType: "transparency", duration: 600 },
        { shapeId: 6, class: "emphasis", emphasisType: "pulse", duration: 500 },
        { shapeId: 7, class: "emphasis", emphasisType: "boldFlash", duration: 500 },
      ],
    },

    // Slide 4: Path animations
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Path Animations" },
            properties: {
              fill: "2F5496",
            },
          },
        },
        {
          shape: {
            id: 2,
            x: "1.3cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "2.1cm",
            textBody: { text: "Line Path" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
        {
          shape: {
            id: 3,
            x: "9.3cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "2.1cm",
            textBody: { text: "Arc Path" },
            properties: {
              fill: "70AD47",
            },
          },
        },
        {
          shape: {
            id: 4,
            x: "1.3cm",
            y: "6.6cm",
            width: "5.3cm",
            height: "2.1cm",
            textBody: { text: "Circle Path" },
            properties: {
              fill: "FFC000",
            },
          },
        },
        {
          shape: {
            id: 5,
            x: "9.3cm",
            y: "6.6cm",
            width: "5.3cm",
            height: "2.1cm",
            textBody: { text: "Custom Path" },
            properties: {
              fill: "5B9BD5",
            },
          },
        },
      ],
      animations: [
        { shapeId: 2, pathType: "line", duration: 1000 },
        { shapeId: 3, pathType: "arc", duration: 1000 },
        { shapeId: 4, pathType: "circle", duration: 1500 },
        {
          shapeId: 5,
          pathType: "customPath",
          path: "M 0 0 L 100 0 L 100 100 L 0 100 Z",
          pathEditMode: "relative",
          duration: 1200,
        },
      ],
    },

    // Slide 5: Extended animations — setBehavior, iterate, command
    {
      children: [
        {
          shape: {
            id: 2,
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Extended: Set / Iterate / Command" },
            properties: {
              fill: "4472C4",
            },
          },
        },
        // setBehavior: instantly set opacity to 0.5, then fade in
        {
          shape: {
            id: 3,
            name: "SetFadeBox",
            x: "1.3cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.1cm",
            textBody: { text: "Set + Fade" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
        // Iterate: animate text per-character
        {
          shape: {
            id: 4,
            x: "1.3cm",
            y: "6.1cm",
            width: "10.6cm",
            height: "2.1cm",
            textBody: { text: "Per-character animation" },
            properties: {
              fill: "70AD47",
            },
          },
        },
        // Command: generic call command
        {
          shape: {
            id: 5,
            x: "1.3cm",
            y: "9.0cm",
            width: "7.9cm",
            height: "2.1cm",
            textBody: { text: "Command" },
            properties: {
              fill: "FFC000",
            },
          },
        },
      ],
      animations: [
        { shapeId: 2, type: "fade", duration: 800 },
        {
          // shapeName resolves to the named shape's cNvPr id at compile time —
          // the fresh-authoring channel (shapeId is the round-trip channel).
          shapeName: "SetFadeBox",
          type: "fade",
          duration: 600,
          setBehavior: {
            attributeName: "style.opacity",
            toValue: "0.5",
            toType: "string",
          },
        },
        {
          shapeId: 4,
          type: "fade",
          duration: 300,
          iterate: {
            type: "element",
            interval: 200,
            backwards: false,
          },
        },
        {
          shapeId: 5,
          commandType: "call",
          command: "animate",
          duration: 1000,
          trigger: "withPrevious",
        },
      ],
    },

    // Slide 6: Build animations, sub-time nodes, end conditions, variant values
    {
      children: [
        {
          shape: {
            id: 2,
            x: "1.3cm",
            y: "0.8cm",
            width: "15.9cm",
            height: "1.6cm",
            textBody: { text: "Build / Sub-Time / Variant Values" },
            properties: {
              fill: "2F5496",
            },
          },
        },
        // Color change with colorFrom/colorTo and colorByRgb
        {
          shape: {
            id: 3,
            x: "1.3cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "2.1cm",
            textBody: { text: "Color From/To" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
        // Motion path with from/rCtr
        {
          shape: {
            id: 4,
            x: "10.6cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "2.1cm",
            textBody: { text: "Motion From/RCtr" },
            properties: {
              fill: "70AD47",
            },
          },
        },
        // Property animation with variant values (boolVal, intVal, fltVal)
        {
          shape: {
            id: 5,
            x: "1.3cm",
            y: "6.3cm",
            width: "7.9cm",
            height: "2.1cm",
            textBody: { text: "Variant Int" },
            properties: {
              fill: "5B9BD5",
            },
          },
        },
        // Animation with endConditions and subTimeNodes
        {
          shape: {
            id: 6,
            x: "10.6cm",
            y: "6.3cm",
            width: "7.9cm",
            height: "2.1cm",
            textBody: { text: "End Cond / SubTN" },
            properties: {
              fill: "BF8F00",
            },
          },
        },
        // Animation with exclusive mode and iterate percentage
        {
          shape: {
            id: 7,
            x: "1.3cm",
            y: "9.5cm",
            width: "7.9cm",
            height: "2.1cm",
            textBody: { text: "Excl + tmPct" },
            properties: {
              fill: "7030A0",
            },
          },
        },
      ],
      animations: [
        {
          shapeId: 2,
          type: "fade",
          duration: 500,
          // Build list: paragraph build for shape 2
          builds: [
            { type: "paragraph", shapeId: 3, groupId: 1, build: "paragraph" },
            { type: "diagram", shapeId: 4, groupId: 2, diagramBuild: "depthByNode" },
          ],
        },
        {
          shapeId: 3,
          // groupId pairs this animation with the bldP entry in the build list
          groupId: 1,
          class: "emphasis",
          emphasisType: "colorChange",
          colorFrom: "FFC000",
          colorTo: "FF0000",
          colorByRgb: { r: "-100%", g: "0%", b: "-100%" },
          duration: 800,
        },
        {
          shapeId: 4,
          // groupId pairs this animation with the bldDgm entry in the build list
          groupId: 2,
          pathType: "arc",
          duration: 1000,
          motionFrom: { x: "0", y: "0" },
          motionRotationCenter: { x: "50000", y: "50000" },
        },
        {
          shapeId: 5,
          attributeName: "style.width",
          calcMode: "linear",
          variantInt: 1,
          from: "0",
          to: "100000",
          duration: 600,
        },
        {
          shapeId: 6,
          type: "fade",
          duration: 500,
          endConditions: [{ delay: "5000" }],
          endSyncCondition: { event: "onNext", delay: "0" },
          subTimeNodes: [{ duration: 300, delay: 100 }],
        },
        {
          shapeId: 7,
          type: "fade",
          duration: 500,
          exclusiveMode: true,
          iterate: { type: "element", iteratePercentage: 50 },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/9-animation.pptx", buffer);
