import type { ColorMappingOptions } from "@office-open/core";
import type { TextListStyleGroupOptions } from "@office-open/core/drawing";
import type { BackgroundOptions } from "@parts/background";
import type { SlideChild } from "@parts/slide/slide-child";

import type { HeaderFooterOptions } from "./handout-master";

export type { HeaderFooterOptions };

/** Options for the notes master (p:notesMaster). */
export interface NotesMasterOptions {
  /** Background (p:bg); defaults to the MS Office bgRef idx="1001". */
  background?: BackgroundOptions;
  /** Custom spTree shapes (p:spTree children after the group header). */
  children?: SlideChild[];
  /** Color mapping overrides (p:clrMap). */
  colorMapping?: Partial<ColorMappingOptions>;
  /** Header/footer settings (p:hf). */
  headerFooter?: HeaderFooterOptions;
  /** Notes text style (p:notesStyle, CT_TextListStyle); defaults to the Office 9-level notes style. */
  notesStyle?: TextListStyleGroupOptions;
}

/** MS Office default notes style — 9 levels, 12pt minor font, incremental margins. */
export const DEFAULT_NOTES_STYLE: TextListStyleGroupOptions = {
  levels: [0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600].map(
    (marginIndent) => ({
      alignment: "l",
      marginIndent,
      defaultTabSize: 914400,
      rtl: false,
      eastAsianLineBreak: true,
      latinLineBreak: false,
      hangingPunctuation: true,
      defaultRun: {
        size: 12,
        kern: 12,
        schemeColor: "tx1",
        latin: "+mn-lt",
        eastAsia: "+mn-ea",
        complexScript: "+mn-cs",
      },
    }),
  ),
};
