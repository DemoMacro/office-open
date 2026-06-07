import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

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
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Entrance Animations" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            id: 3,
            x: 50,
            y: 120,
            width: 300,
            height: 100,
            textBody: { text: "Appear" },
            fill: "ED7D31",
          },
        },
        {
          shape: {
            id: 4,
            x: 400,
            y: 120,
            width: 300,
            height: 100,
            textBody: { text: "Fly In (from left)" },
            fill: "70AD47",
          },
        },
        {
          shape: {
            id: 5,
            x: 50,
            y: 250,
            width: 300,
            height: 100,
            textBody: { text: "Wipe (down)" },
            fill: "FFC000",
          },
        },
        {
          shape: {
            id: 6,
            x: 400,
            y: 250,
            width: 300,
            height: 100,
            textBody: { text: "Dissolve" },
            fill: "5B9BD5",
          },
        },
        {
          shape: {
            id: 7,
            x: 50,
            y: 380,
            width: 300,
            height: 100,
            textBody: { text: "Zoom In" },
            fill: "BF8F00",
          },
        },
        {
          shape: {
            id: 8,
            x: 400,
            y: 380,
            width: 300,
            height: 100,
            textBody: { text: "Split (horizontal)" },
            fill: "7030A0",
          },
        },
      ],
      animations: [
        { shapeId: 2, options: { type: "fade", duration: 800 } },
        { shapeId: 3, options: { type: "appear" } },
        { shapeId: 4, options: { type: "fly", direction: "left", duration: 600 } },
        { shapeId: 5, options: { type: "wipe", direction: "down" } },
        { shapeId: 6, options: { type: "dissolve", duration: 1000 } },
        { shapeId: 7, options: { type: "zoom" } },
        { shapeId: 8, options: { type: "split", direction: "horizontal" } },
      ],
    },

    // Slide 2: Exit animations
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Exit Animations" },
            fill: "C00000",
          },
        },
        {
          shape: {
            id: 2,
            x: 50,
            y: 120,
            width: 300,
            height: 100,
            textBody: { text: "Fade Out" },
            fill: "ED7D31",
          },
        },
        {
          shape: {
            id: 3,
            x: 400,
            y: 120,
            width: 300,
            height: 100,
            textBody: { text: "Fly Out (right)" },
            fill: "70AD47",
          },
        },
        {
          shape: {
            id: 4,
            x: 50,
            y: 250,
            width: 300,
            height: 100,
            textBody: { text: "Wipe Out (up)" },
            fill: "FFC000",
          },
        },
        {
          shape: {
            id: 5,
            x: 400,
            y: 250,
            width: 300,
            height: 100,
            textBody: { text: "Dissolve Out" },
            fill: "5B9BD5",
          },
        },
        {
          shape: {
            id: 6,
            x: 50,
            y: 380,
            width: 300,
            height: 100,
            textBody: { text: "Zoom Out" },
            fill: "BF8F00",
          },
        },
        {
          shape: {
            id: 7,
            x: 400,
            y: 380,
            width: 300,
            height: 100,
            textBody: { text: "Split Out" },
            fill: "7030A0",
          },
        },
      ],
      animations: [
        { shapeId: 2, options: { type: "fade", class: "exit", duration: 800 } },
        { shapeId: 3, options: { type: "fly", class: "exit", direction: "right", duration: 600 } },
        { shapeId: 4, options: { type: "wipe", class: "exit", direction: "up" } },
        { shapeId: 5, options: { type: "dissolve", class: "exit", duration: 1000 } },
        { shapeId: 6, options: { type: "zoom", class: "exit" } },
        { shapeId: 7, options: { type: "split", class: "exit", direction: "horizontal" } },
      ],
    },

    // Slide 3: Emphasis animations
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Emphasis Animations" },
            fill: "548235",
          },
        },
        {
          shape: {
            id: 2,
            x: 50,
            y: 120,
            width: 300,
            height: 100,
            textBody: { text: "Grow/Shrink" },
            fill: "ED7D31",
          },
        },
        {
          shape: {
            id: 3,
            x: 400,
            y: 120,
            width: 300,
            height: 100,
            textBody: { text: "Spin" },
            fill: "70AD47",
          },
        },
        {
          shape: {
            id: 4,
            x: 50,
            y: 250,
            width: 300,
            height: 100,
            textBody: { text: "Color Change" },
            fill: "FFC000",
          },
        },
        {
          shape: {
            id: 5,
            x: 400,
            y: 250,
            width: 300,
            height: 100,
            textBody: { text: "Transparency" },
            fill: "5B9BD5",
          },
        },
        {
          shape: {
            id: 6,
            x: 50,
            y: 380,
            width: 300,
            height: 100,
            textBody: { text: "Pulse" },
            fill: "BF8F00",
          },
        },
        {
          shape: {
            id: 7,
            x: 400,
            y: 380,
            width: 300,
            height: 100,
            textBody: { text: "Bold Flash" },
            fill: "7030A0",
          },
        },
      ],
      animations: [
        { shapeId: 2, options: { class: "emph", emphasisType: "growShrink", duration: 800 } },
        { shapeId: 3, options: { class: "emph", emphasisType: "spin", duration: 1000 } },
        {
          shapeId: 4,
          options: {
            class: "emph",
            emphasisType: "colorChange",
            color: "FF0000",
            duration: 800,
          },
        },
        { shapeId: 5, options: { class: "emph", emphasisType: "transparency", duration: 600 } },
        { shapeId: 6, options: { class: "emph", emphasisType: "pulse", duration: 500 } },
        { shapeId: 7, options: { class: "emph", emphasisType: "boldFlash", duration: 500 } },
      ],
    },

    // Slide 4: Path animations
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Path Animations" },
            fill: "2F5496",
          },
        },
        {
          shape: {
            id: 2,
            x: 50,
            y: 120,
            width: 200,
            height: 80,
            textBody: { text: "Line Path" },
            fill: "ED7D31",
          },
        },
        {
          shape: {
            id: 3,
            x: 350,
            y: 120,
            width: 200,
            height: 80,
            textBody: { text: "Arc Path" },
            fill: "70AD47",
          },
        },
        {
          shape: {
            id: 4,
            x: 50,
            y: 250,
            width: 200,
            height: 80,
            textBody: { text: "Circle Path" },
            fill: "FFC000",
          },
        },
        {
          shape: {
            id: 5,
            x: 350,
            y: 250,
            width: 200,
            height: 80,
            textBody: { text: "Custom Path" },
            fill: "5B9BD5",
          },
        },
      ],
      animations: [
        { shapeId: 2, options: { pathType: "line", duration: 1000 } },
        { shapeId: 3, options: { pathType: "arc", duration: 1000 } },
        { shapeId: 4, options: { pathType: "circle", duration: 1500 } },
        {
          shapeId: 5,
          options: {
            pathType: "customPath",
            path: "M 0 0 L 100 0 L 100 100 L 0 100 Z",
            duration: 1200,
          },
        },
      ],
    },

    // Slide 5: Extended animations — setBehavior, iterate, command
    {
      children: [
        {
          shape: {
            id: 2,
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Extended: Set / Iterate / Command" },
            fill: "4472C4",
          },
        },
        // setBehavior: instantly set opacity to 0.5, then fade in
        {
          shape: {
            id: 3,
            x: 50,
            y: 120,
            width: 300,
            height: 80,
            textBody: { text: "Set + Fade" },
            fill: "ED7D31",
          },
        },
        // Iterate: animate text per-character
        {
          shape: {
            id: 4,
            x: 50,
            y: 230,
            width: 400,
            height: 80,
            textBody: { text: "Per-character animation" },
            fill: "70AD47",
          },
        },
        // Command: generic call command
        {
          shape: {
            id: 5,
            x: 50,
            y: 340,
            width: 300,
            height: 80,
            textBody: { text: "Command" },
            fill: "FFC000",
          },
        },
      ],
      animations: [
        { shapeId: 2, options: { type: "fade", duration: 800 } },
        {
          shapeId: 3,
          options: {
            type: "fade",
            duration: 600,
            setBehavior: {
              attributeName: "style.opacity",
              toValue: "0.5",
              toType: "string",
            },
          },
        },
        {
          shapeId: 4,
          options: {
            type: "fade",
            duration: 300,
            iterate: {
              type: "el",
              interval: 200,
              backwards: false,
            },
          },
        },
        {
          shapeId: 5,
          options: {
            commandType: "call",
            command: "animate",
            duration: 1000,
            trigger: "withPrevious",
          },
        },
      ],
    },

    // Slide 6: Build animations, sub-time nodes, end conditions, variant values
    {
      children: [
        {
          shape: {
            id: 2,
            x: 50,
            y: 30,
            width: 600,
            height: 60,
            textBody: { text: "Build / Sub-Time / Variant Values" },
            fill: "2F5496",
          },
        },
        // Color change with colorFrom/colorTo and colorByRgb
        {
          shape: {
            id: 3,
            x: 50,
            y: 120,
            width: 300,
            height: 80,
            textBody: { text: "Color From/To" },
            fill: "ED7D31",
          },
        },
        // Motion path with from/rCtr
        {
          shape: {
            id: 4,
            x: 400,
            y: 120,
            width: 200,
            height: 80,
            textBody: { text: "Motion From/RCtr" },
            fill: "70AD47",
          },
        },
        // Property animation with variant values (boolVal, intVal, fltVal)
        {
          shape: {
            id: 5,
            x: 50,
            y: 240,
            width: 300,
            height: 80,
            textBody: { text: "Variant Int" },
            fill: "5B9BD5",
          },
        },
        // Animation with endConditions and subTimeNodes
        {
          shape: {
            id: 6,
            x: 400,
            y: 240,
            width: 300,
            height: 80,
            textBody: { text: "End Cond / SubTN" },
            fill: "BF8F00",
          },
        },
        // Animation with exclusive mode and iterate percentage
        {
          shape: {
            id: 7,
            x: 50,
            y: 360,
            width: 300,
            height: 80,
            textBody: { text: "Excl + tmPct" },
            fill: "7030A0",
          },
        },
      ],
      animations: [
        {
          shapeId: 2,
          options: {
            type: "fade",
            duration: 500,
            // Build list: paragraph build for shape 2
            builds: [
              { type: "paragraph", spid: 3, grpId: 1, build: "p" },
              { type: "diagram", spid: 4, grpId: 2, diagramBuild: "depthByNode" },
            ],
          },
        },
        {
          shapeId: 3,
          options: {
            class: "emph",
            emphasisType: "colorChange",
            colorFrom: "FFC000",
            colorTo: "FF0000",
            colorByRgb: { r: "-100%", g: "0%", b: "-100%" },
            duration: 800,
          },
        },
        {
          shapeId: 4,
          options: {
            pathType: "arc",
            duration: 1000,
            motionFrom: { x: "0", y: "0" },
            motionRotationCenter: { x: "50000", y: "50000" },
          },
        },
        {
          shapeId: 5,
          options: {
            attributeName: "style.width",
            calcMode: "lin",
            variantInt: 1,
            from: "0",
            to: "100000",
            duration: 600,
          },
        },
        {
          shapeId: 6,
          options: {
            type: "fade",
            duration: 500,
            endConditions: [{ delay: "5000" }],
            endSyncCondition: { event: "onNext", delay: "0" },
            subTimeNodes: [{ duration: 300, delay: 100 }],
          },
        },
        {
          shapeId: 7,
          options: {
            type: "fade",
            duration: 500,
            exclusiveMode: true,
            iterate: { type: "el", iteratePercentage: 50000 },
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
