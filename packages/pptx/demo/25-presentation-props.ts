// Presentation properties: presPr (web/print/show), viewPr, photoAlbum, modifyVerifier.

import * as fs from "fs";

import { Presentation, Shape, Packer, Paragraph, TextRun } from "@office-open/pptx";

const pres = new Presentation({
  title: "Presentation Properties Demo",
  creator: "Demo",

  // Presentation root attributes
  firstSlideNum: 5,
  serverZoom: "334000",
  rtl: false,
  autoCompressPictures: true,
  bookmarkIdSeed: 1234,

  // Photo album
  photoAlbum: {
    blackWhite: false,
    showCaptions: true,
    layout: "2pic",
    frame: "frameStyle1",
  },

  // Modify verifier (encryption)
  modifyVerifier: {
    password: "secret",
    spinValue: 100000,
    cryptoProviderType: "rsaAES",
    cryptoAlgorithmClass: "hash",
    cryptoAlgorithmType: "typeAny",
    cryptoAlgorithmSid: 12,
  },

  // View properties
  view: {
    lastView: "slideView",
    showComments: true,
    gridSpacing: { cx: 50800, cy: 50800 },
    normalView: {
      showOutlineIcons: true,
      snapVertSplitter: false,
      vertBarState: "restored",
      horzBarState: "restored",
      preferSingleView: false,
    },
    slideView: {
      snapToGrid: true,
      snapToObjects: true,
      showGuides: true,
      varScale: true,
    },
  },

  // Presentation properties (presPr): web, print, show
  show: {
    type: "kiosk",
    restart: 300000,
    showNarration: true,
    showAnimation: true,
    useTimings: true,
    penColor: "FF0000",
  },
  // Web properties (rendered in presPr)
  web: {
    showAnimation: true,
    resizeGraphics: true,
    allowPng: false,
    relyOnVml: false,
    organizeInFolders: true,
    useLongFilenames: true,
  },
  // Print properties (rendered in presPr)
  print: {
    printWhat: "handouts4",
    colorMode: "color",
    hiddenSlides: false,
    scaleToFitPaper: true,
    frameSlides: true,
  },
  // HTML publish properties (rendered in presPr)
  htmlPublish: {
    showSpeakerNotes: true,
    title: "Published Slides",
    rId: "rId1",
  },

  slides: [
    {
      children: [
        new Shape({
          x: 100,
          y: 100,
          width: 600,
          height: 200,
          textBody: {
            text: "Presentation Properties Demo",
          },
        }),
      ],
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 100,
          width: 600,
          height: 200,
          textBody: {
            children: [new Paragraph({ children: [new TextRun("Slide with full props")] })],
          },
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
