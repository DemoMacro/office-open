import type { UniversalMeasure } from "@office-open/core";
import { SP_TREE_HEADER } from "@shared/constants";

import {
  buildColorMapAttrs,
  buildHfAttrs,
  type ColorMapOptions,
  type HeaderFooterOptions,
} from "./handout-master";

export type { ColorMapOptions, HeaderFooterOptions };

/** Notes style level override */
export interface NotesLevelProperties {
  /** Font size in hundredths of a point (e.g., 1200 = 12pt), ST_TextPoint */
  fontSize?: number | UniversalMeasure;
  /** Left margin (ST_Coordinate32) */
  marginLeft?: number | UniversalMeasure;
  /** Alignment ("l" | "ctr" | "r" | "just") */
  alignment?: string;
}

/** Options for notes master parameterization */
export interface NotesMasterOptions {
  /** Color map overrides */
  colorMap?: ColorMapOptions;
  /** Header/footer settings */
  headerFooter?: HeaderFooterOptions;
  /** Notes style overrides (levels 1-9) */
  notesStyle?: NotesLevelProperties[];
}

const DEFAULT_LEVEL_MARGINS = [
  0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600,
];

function buildNotesStyleXml(levels?: NotesLevelProperties[]): string {
  const parts: string[] = ["<p:notesStyle>"];
  for (let i = 0; i < 9; i++) {
    const level = levels?.[i];
    const marL = level?.marginLeft ?? DEFAULT_LEVEL_MARGINS[i];
    const algn = level?.alignment ?? "l";
    const sz = level?.fontSize ?? 1200;
    parts.push(
      `<a:lvl${i + 1}pPr marL="${marL}" algn="${algn}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
        `<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` +
        `<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`,
    );
  }
  parts.push("</p:notesStyle>");
  return parts.join("");
}

export function buildNotesMasterXml(options?: NotesMasterOptions): string {
  const colorMap = buildColorMapAttrs(options?.colorMap);
  const hf = buildHfAttrs(options?.headerFooter);
  const notesStyle = buildNotesStyleXml(options?.notesStyle);
  return (
    '<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    `<p:spTree>${SP_TREE_HEADER}</p:spTree></p:cSld>` +
    `<p:clrMap ${colorMap}/>` +
    `<p:hf ${hf}/>` +
    notesStyle +
    "</p:notesMaster>"
  );
}
