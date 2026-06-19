/**
 * Content Types module for PPTX packages.
 *
 * Pure data builder — no XmlComponent inheritance.
 *
 * @module
 */
import { escapeXml } from "@office-open/xml";

const PPTX_MAIN =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const PPTX_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const PPTX_SLIDE_MASTER =
  "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
const PPTX_SLIDE_LAYOUT =
  "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const PPTX_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
const PPTX_NOTES = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const PPTX_NOTES_MASTER =
  "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
const PPTX_HANDOUT_MASTER =
  "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";
const PPTX_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
const PPTX_COMMENT_AUTHORS =
  "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
const PPTX_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const PPTX_PRES_PROPS =
  "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
const PPTX_VIEW_PROPS =
  "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
const PPTX_TABLE_STYLES =
  "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
const PPTX_SLIDE_SYNC =
  "application/vnd.openxmlformats-officedocument.presentationml.slideSyncProperties+xml";
const CUSTOM_PROPS = "application/vnd.openxmlformats-officedocument.custom-properties+xml";

type EntryType = "Default" | "Override";

interface ContentEntry {
  type: EntryType;
  contentType: string;
  key: string;
}

const STATIC_ENTRIES: ContentEntry[] = [
  {
    type: "Default",
    contentType: "application/vnd.openxmlformats-package.relationships+xml",
    key: "rels",
  },
  { type: "Default", contentType: "application/xml", key: "xml" },
  { type: "Default", contentType: "image/png", key: "png" },
  { type: "Default", contentType: "image/jpeg", key: "jpeg" },
  { type: "Default", contentType: "image/jpeg", key: "jpg" },
  { type: "Default", contentType: "video/mp4", key: "mp4" },
  { type: "Override", contentType: PPTX_MAIN, key: "/ppt/presentation.xml" },
  {
    type: "Override",
    contentType: "application/vnd.openxmlformats-package.core-properties+xml",
    key: "/docProps/core.xml",
  },
  {
    type: "Override",
    contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    key: "/docProps/app.xml",
  },
  { type: "Override", contentType: PPTX_THEME, key: "/ppt/theme/theme1.xml" },
  { type: "Override", contentType: PPTX_PRES_PROPS, key: "/ppt/presProps.xml" },
  { type: "Override", contentType: PPTX_VIEW_PROPS, key: "/ppt/viewProps.xml" },
  { type: "Override", contentType: PPTX_TABLE_STYLES, key: "/ppt/tableStyles.xml" },
];

function formatEntry(e: ContentEntry): string {
  if (e.type === "Default") {
    return `<Default ContentType="${escapeXml(e.contentType)}" Extension="${escapeXml(e.key)}"/>`;
  }
  return `<Override ContentType="${escapeXml(e.contentType)}" PartName="${escapeXml(e.key)}"/>`;
}

const STATIC_XML =
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  STATIC_ENTRIES.map(formatEntry).join("");

/** Pure data builder for [Content_Types].xml entries. */
export class ContentTypes {
  private entries: ContentEntry[] = [];

  public addSlide(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_SLIDE,
      key: `/ppt/slides/slide${index}.xml`,
    });
  }

  public addNotesSlide(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_NOTES,
      key: `/ppt/notesSlides/notesSlide${index}.xml`,
    });
  }

  public addNotesMaster(): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_NOTES_MASTER,
      key: "/ppt/notesMasters/notesMaster1.xml",
    });
  }

  public addHandoutMaster(): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_HANDOUT_MASTER,
      key: "/ppt/handoutMasters/handoutMaster1.xml",
    });
  }

  public addComments(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_COMMENTS,
      key: `/ppt/comments/comment${index}.xml`,
    });
  }

  public addCommentAuthors(): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_COMMENT_AUTHORS,
      key: "/ppt/commentAuthors.xml",
    });
  }

  public addChart(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_CHART,
      key: `/ppt/charts/chart${index}.xml`,
    });
  }

  public addDiagramData(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
      key: `/ppt/diagrams/data${index}.xml`,
    });
  }

  public addDiagramLayout(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
      key: `/ppt/diagrams/layout${index}.xml`,
    });
  }

  public addDiagramStyle(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
      key: `/ppt/diagrams/quickStyle${index}.xml`,
    });
  }

  public addDiagramColors(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
      key: `/ppt/diagrams/colors${index}.xml`,
    });
  }

  public addDiagramDrawing(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
      key: `/ppt/diagrams/drawing${index}.xml`,
    });
  }

  public addSlideLayout(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_SLIDE_LAYOUT,
      key: `/ppt/slideLayouts/slideLayout${index}.xml`,
    });
  }

  public addSlideMaster(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_SLIDE_MASTER,
      key: `/ppt/slideMasters/slideMaster${index}.xml`,
    });
  }

  public addCustomProperties(): void {
    this.entries.push({
      type: "Override",
      contentType: CUSTOM_PROPS,
      key: "/docProps/custom.xml",
    });
  }

  public addSlideSyncPr(index: number): void {
    this.entries.push({
      type: "Override",
      contentType: PPTX_SLIDE_SYNC,
      key: `/ppt/slideSyncPr/slideSyncPr${index}.xml`,
    });
  }

  public addTheme(index: number): void {
    if (index !== 1) {
      this.entries.push({
        type: "Override",
        contentType: PPTX_THEME,
        key: `/ppt/theme/theme${index}.xml`,
      });
    }
  }

  /** Serialize to [Content_Types].xml string. */
  public serialize(): string {
    const parts: string[] = [STATIC_XML];
    for (const e of this.entries) {
      parts.push(formatEntry(e));
    }
    parts.push("</Types>");
    return parts.join("");
  }
}
