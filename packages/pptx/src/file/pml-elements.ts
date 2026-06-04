/**
 * PresentationML (p:) element name constants.
 *
 * Covers all pml.xsd elements that are not referenced by existing component classes.
 * Used for XSD coverage tracking — each constant holds the literal element name string.
 *
 * @module
 */

// ── Transition child elements ──

/** p:blinds — Blinds slide transition */
export const BLINDS = "p:blinds";
/** p:checker — Checker slide transition */
export const CHECKER = "p:checker";
/** p:circle — Circle slide transition */
export const CIRCLE = "p:circle";
/** p:comb — Comb slide transition */
export const COMB = "p:comb";
/** p:cover — Cover slide transition */
export const COVER = "p:cover";
/** p:cut — Cut slide transition */
export const CUT = "p:cut";
/** p:diamond — Diamond slide transition */
export const DIAMOND = "p:diamond";
/** p:dissolve — Dissolve slide transition */
export const DISSOLVE = "p:dissolve";
/** p:fade — Fade slide transition */
export const FADE = "p:fade";
/** p:newsflash — Newsflash slide transition */
export const NEWSFLASH = "p:newsflash";
/** p:plus — Plus slide transition */
export const PLUS = "p:plus";
/** p:pull — Pull slide transition */
export const PULL = "p:pull";
/** p:push — Push slide transition */
export const PUSH = "p:push";
/** p:random — Random slide transition */
export const RANDOM = "p:random";
/** p:randomBar — Random bar slide transition */
export const RANDOM_BAR = "p:randomBar";
/** p:split — Split slide transition */
export const SPLIT = "p:split";
/** p:strips — Strips slide transition */
export const STRIPS = "p:strips";
/** p:wedge — Wedge slide transition */
export const WEDGE = "p:wedge";
/** p:wheel — Wheel slide transition */
export const WHEEL = "p:wheel";
/** p:wipe — Wipe slide transition */
export const WIPE = "p:wipe";
/** p:zoom — Zoom slide transition */
export const ZOOM = "p:zoom";

// ── Transition sound elements ──

/** p:stSnd — Start sound action for transition */
export const ST_SND = "p:stSnd";
/** p:endSnd — End sound action for transition */
export const END_SND = "p:endSnd";
/** p:sndAc — Sound action container for transition */
export const SND_AC = "p:sndAc";

// ── Animation — build elements ──

/** p:bldLst — Build list for slide animations */
export const BLD_LST = "p:bldLst";
/** p:bldP — Build paragraph animation */
export const BLD_P = "p:bldP";
/** p:bldDgm — Build diagram animation */
export const BLD_DGM = "p:bldDgm";
/** p:bldOleChart — Build OLE chart animation */
export const BLD_OLE_CHART = "p:bldOleChart";
/** p:bldGraphic — Build graphic element animation */
export const BLD_GRAPHIC = "p:bldGraphic";
/** p:bldAsOne — Build as one animation */
export const BLD_AS_ONE = "p:bldAsOne";
/** p:bldSub — Build sub-element animation */
export const BLD_SUB = "p:bldSub";

// ── Animation — timing ──

/** p:tn — Timeline node */
export const TN = "p:tn";
/** p:subTnLst — Sub-timeline node list */
export const SUB_TN_LST = "p:subTnLst";
/** p:endCondLst — End condition list */
export const END_COND_LST = "p:endCondLst";
/** p:endSync — End synchronization */
export const END_SYNC = "p:endSync";
/** p:tmpl — Animation template */
export const TMPL = "p:tmpl";
/** p:tmplLst — Animation template list */
export const TMPL_LST = "p:tmplLst";
/** p:tmPct — Timeline percentage */
export const TM_PCT = "p:tmPct";
/** p:excl — Exclusive timing node */
export const EXCL = "p:excl";

// ── Animation — targets ──

/** p:inkTgt — Ink target element */
export const INK_TGT = "p:inkTgt";
/** p:sndTgt — Sound target element */
export const SND_TGT = "p:sndTgt";
/** p:spTgt sub-elements — sub-shape target */
export const SUB_SP = "p:subSp";
/** p:graphicEl — Graphic element target */
export const GRAPHIC_EL = "p:graphicEl";
/** p:oleChartEl — OLE chart element target */
export const OLE_CHART_EL = "p:oleChartEl";
/** p:rCtr — Runtime context */
export const R_CTR = "p:rCtr";

// ── Animation — value types ──

/** p:boolVal — Boolean animation value */
export const BOOL_VAL = "p:boolVal";
/** p:intVal — Integer animation value */
export const INT_VAL = "p:intVal";
/** p:fltVal — Float animation value */
export const FLT_VAL = "p:fltVal";
/** p:strVal note: already referenced in timing.ts, included for completeness */
export const STR_VAL = "p:strVal";
/** p:clrVal — Color animation value */
export const CLR_VAL = "p:clrVal";

// ── Animation — from/to ──

/** p:from — Animation from value */
export const FROM = "p:from";

// ── Animation — color components ──

/** p:rgb — RGB color value for animation */
export const RGB = "p:rgb";
/** p:hsl — HSL color value for animation */
export const HSL = "p:hsl";

// ── Animation — direction/motion ──

/** p:rtn — Rotation transform */
export const RTN = "p:rtn";

// ── Build — text styles ──

/** p:bold — Bold text build style */
export const BOLD = "p:bold";
/** p:italic — Italic text build style */
export const ITALIC = "p:italic";
/** p:boldItalic — Bold-italic text build style */
export const BOLD_ITALIC = "p:boldItalic";
/** p:regular — Regular text build style */
export const REGULAR = "p:regular";
/** p:style — Text style reference */
export const STYLE = "p:style";
/** p:font — Font element */
export const FONT = "p:font";

// ── Slide content ──

/** p:contentPart — Content part reference */
export const CONTENT_PART = "p:contentPart";
/** p:sldLst — Slide list */
export const SLD_LST = "p:sldLst";
/** p:photoAlbum — Photo album settings */
export const PHOTO_ALBUM = "p:photoAlbum";

// ── Controls ──

/** p:control — Embedded control */
export const CONTROL = "p:control";
/** p:controls — Controls container */
export const CONTROLS = "p:controls";

// ── Custom data ──

/** p:custData — Custom data element */
export const CUST_DATA = "p:custData";
/** p:custDataLst — Custom data list */
export const CUST_DATA_LST = "p:custDataLst";
/** p:custShow — Custom show */
export const CUST_SHOW = "p:custShow";
/** p:custShowLst — Custom show list */
export const CUST_SHOW_LST = "p:custShowLst";

// ── Tags ──

/** p:tag — Tag element */
export const TAG = "p:tag";
/** p:tagLst — Tag list */
export const TAG_LST = "p:tagLst";
/** p:tags — Tags container */
export const TAGS = "p:tags";

// ── Embedded fonts ──

/** p:embeddedFont — Embedded font */
export const EMBEDDED_FONT = "p:embeddedFont";
/** p:embeddedFontLst — Embedded font list */
export const EMBEDDED_FONT_LST = "p:embeddedFontLst";

// ── Presentation properties ──

/** p:htmlPubPr — HTML publish properties */
export const HTML_PUB_PR = "p:htmlPubPr";
/** p:webPr — Web properties */
export const WEB_PR = "p:webPr";
/** p:prnPr — Print properties */
export const PRN_PR = "p:prnPr";
/** p:modifyVerifier — Modification verifier */
export const MODIFY_VERIFIER = "p:modifyVerifier";
/** p:smartTags — Smart tags */
export const SMART_TAGS = "p:smartTags";
/** p:sldSyncPr — Slide synchronization properties */
export const SLD_SYNC_PR = "p:sldSyncPr";

// ── View properties ──

/** p:notesViewPr — Notes view properties */
export const NOTES_VIEW_PR = "p:notesViewPr";
/** p:outlineViewPr — Outline view properties */
export const OUTLINE_VIEW_PR = "p:outlineViewPr";
/** p:sorterViewPr — Sorter view properties */
export const SORTER_VIEW_PR = "p:sorterViewPr";

// ── Color ──

/** p:clrMru — Color most-recently-used */
export const CLR_MRU = "p:clrMru";

// ── Guide ──

/** p:guide — Drawing guide */
export const GUIDE = "p:guide";

// ── Text ──

/** p:kinsoku — Kinsoku line breaking settings */
export const KINSOKU = "p:kinsoku";

// ── Progress ──

/** p:progress — Progress element */
export const PROGRESS = "p:progress";
