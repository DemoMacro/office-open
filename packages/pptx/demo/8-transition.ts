import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Minimal 16-bit mono 8 kHz WAV: 44-byte header + 800 samples of silence.
function makeSilenceWav(): Uint8Array {
  const samples = 800;
  const data = new Uint8Array(44 + samples * 2);
  const view = new DataView(data.buffer);
  const ascii = (offset: number, text: string) => {
    for (let i = 0; i < text.length; i++) view.setUint8(offset + i, text.charCodeAt(i));
  };
  ascii(0, "RIFF");
  view.setUint32(4, 36 + samples * 2, true);
  ascii(8, "WAVE");
  ascii(12, "fmt ");
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true); // PCM
  view.setUint16(22, 1, true); // mono
  view.setUint32(24, 8000, true);
  view.setUint32(28, 16000, true);
  view.setUint16(32, 2, true);
  view.setUint16(34, 16, true);
  ascii(36, "data");
  view.setUint32(40, samples * 2, true);
  return data;
}

const options: PresentationOptions = {
  title: "Transition Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Fade Transition" },
            fill: "4472C4",
          },
        },
      ],
      transition: { type: "fade", speed: "medium" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Push Transition (Right)" },
            fill: "ED7D31",
          },
        },
      ],
      transition: { type: "push", direction: "right", speed: "slow" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Wipe Transition (Down)" },
            fill: "70AD47",
          },
        },
      ],
      transition: { type: "wipe", direction: "down" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Cover Transition (From Right)" },
            fill: "FFC000",
          },
        },
      ],
      transition: { type: "cover", direction: "right" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Split Transition" },
            fill: "5B9BD5",
          },
        },
      ],
      transition: { type: "split", orient: "horz", direction: "out" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Wheel Transition (4 spokes)" },
            fill: "BF8F00",
          },
        },
      ],
      transition: { type: "wheel", spokes: 4 },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Dissolve Transition" },
            fill: "7030A0",
          },
        },
      ],
      transition: { type: "dissolve", speed: "slow" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Random Transition" },
            fill: "C00000",
          },
        },
      ],
      transition: { type: "random" },
    },
    // Transition with start sound
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Fade with Sound" },
            fill: "4472C4",
          },
        },
      ],
      transition: {
        type: "fade",
        speed: "medium",
        startSound: { data: makeSilenceWav(), type: "wav", name: "silence", loop: true },
      },
    },
    // Transition with stop previous sound
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Push (Stop Sound)" },
            fill: "ED7D31",
          },
        },
      ],
      transition: { type: "push", direction: "right", stopPreviousSound: true },
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/8-transition.pptx", buffer);
