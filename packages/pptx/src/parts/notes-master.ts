import type { ColorMappingOptions } from "@office-open/core";
import type { TextListStyleOptions } from "@office-open/core/drawing";
import type { ThemeOptions } from "@office-open/core/theme";
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
  notesStyle?: TextListStyleOptions;
  /** The notes master's own theme part (ppt/theme/themeN.xml); defaults to the Office theme. */
  theme?: ThemeOptions;
}

/** MS Office default notes style — 9 levels, 12pt minor font, incremental margins. */
export const DEFAULT_NOTES_STYLE: TextListStyleOptions = {
  levels: [0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600].map(
    (marginIndent) => ({
      alignment: "left",
      marginIndent,
      defTabSize: 914400,
      rightToLeft: false,
      eastAsianLineBreak: true,
      latinLineBreak: false,
      hangingPunctuation: true,
      defaultRunProperties: {
        size: 12,
        kern: 12,
        fill: { type: "solid", color: { value: "tx1" } },
        font: { latin: "+mn-lt", eastAsia: "+mn-ea", complexScript: "+mn-cs" },
      },
    }),
  ),
};
