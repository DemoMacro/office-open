import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Comment Demo",
  creator: "Demo",
  slides: [
    // Slide 1: Two comments from different authors
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Slide with Comments" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: 50,
            y: 150,
            width: 500,
            height: 60,
            textBody: { text: "Content to review" },
            fill: "ED7D31",
          },
        },
      ],
      comments: [
        {
          author: "Alice Wang",
          text: "The title looks great!",
          x: 200,
          y: 50,
          date: "2026-05-15T10:00:00Z",
        },
        {
          author: "Bob Li",
          text: "Consider changing the content color.",
          x: 300,
          y: 180,
          initials: "BL",
          date: "2026-05-15T11:00:00Z",
        },
      ],
    },

    // Slide 2: Comment from existing author (tests author reuse)
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Another Slide" },
            fill: "70AD47",
          },
        },
      ],
      comments: [
        {
          author: "Alice Wang",
          text: "Please review this slide carefully.",
          x: 200,
          y: 50,
          date: "2026-05-15T12:00:00Z",
        },
      ],
    },

    // Slide 3: No comments
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "No Comments Here" },
            fill: "FFC000",
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
