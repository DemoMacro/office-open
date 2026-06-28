// Presentation properties: presPr (web/print/show), viewPr, photoAlbum, modifyVerifier.

import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
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

  // Modify verifier (write protection)
  // cryptAlgorithmSid: 14 = SHA-512 (12 = SHA-256, 13 = SHA-384)
  modifyVerifier: {
    password: "secret",
    cryptoProviderType: "rsaAES",
    cryptoAlgorithmClass: "hash",
    cryptoAlgorithmType: "typeAny",
    cryptoAlgorithmSid: 14,
  },

  // Embedded fonts — font definition without embedding (rId omitted, valid per XSD minOccurs=0)
  // To embed actual font data, provide regular/bold/italic/boldItalic rId references
  // pointing to real font part relationships (similar to docx font embedding).
  embeddedFonts: [
    {
      font: { typeface: "Calibri", panose: "020F0502020204030204", pitchFamily: 34, charset: 0 },
    },
  ],

  // Custom shows — rId values reference slide relationship IDs from presentation.xml.rels
  // With 1 master: rId1=master, rId2=slide1, rId3=slide2, rId4=presProps, ...
  customShows: [{ name: "Quick Show", id: 1, slides: [{ rId: "rId2" }, { rId: "rId3" }] }],

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
    // Guides
    guides: [
      { orient: "horz", pos: 3429000 },
      { orient: "vert", pos: 4572000 },
    ],
    // Outline view — rId values reference slides via auto-generated viewProps.xml.rels
    outlineView: {
      slides: [{ rId: "rId1" }, { rId: "rId2", collapse: true }],
    },
    // Sorter view
    sorterView: {
      showFormatting: true,
    },
    // Notes view
    notesView: true,
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
  // Color MRU (recent colors)
  colorMru: ["FF0000", "00FF00", "0000FF"],

  slides: [
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "5.3cm",
            textBody: {
              text: "Presentation Properties Demo",
            },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "5.3cm",
            textBody: {
              children: [{ children: [{ text: "Slide with full props" }] }],
            },
          },
        },
      ],
      // Slide sync properties (standalone part, not inside p:sld per transitional XSD)
      slideSync: {
        serverSldId: "abc123",
        serverSldModifiedTime: "2025-01-01T00:00:00Z",
        clientInsertedTime: "2025-01-02T00:00:00Z",
      },
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
