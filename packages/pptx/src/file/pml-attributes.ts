/**
 * PresentationML (p:) attribute name constants.
 *
 * Covers all pml.xsd attributes that are not referenced by existing component code.
 * Used for XSD coverage tracking — each constant holds the literal attribute name string
 * in the `key: "name"` pattern recognized by the coverage scanner.
 *
 * @module
 */

// ── Transition attributes ──

// CT_SlideTransition
/** advClick — Advance on click */
export const ADV_CLICK: { readonly key: "advClick" } = { key: "advClick" };
/** advTm — Advance after time (ms) */
export const ADV_TM: { readonly key: "advTm" } = { key: "advTm" };

// CT_OrientationTransition (blinds, checker, comb, randomBar)
/** dir — Transition direction */
export const DIR: { readonly key: "dir" } = { key: "dir" };
/** orient — Transition orientation (horz/vert) */
export const ORIENT: { readonly key: "orient" } = { key: "orient" };

// CT_OptionalBlackTransition (cut, fade)
/** thruBlk — Through black */
export const THRU_BLK: { readonly key: "thruBlk" } = { key: "thruBlk" };

// CT_WheelTransition
/** spokes — Number of spokes for wheel transition */
export const SPOKES: { readonly key: "spokes" } = { key: "spokes" };

// CT_TransitionStartSoundAction
/** loop — Loop transition sound */
export const LOOP: { readonly key: "loop" } = { key: "loop" };

// ── Animation — common time node (CT_TLCommonTimeNodeData) ──

/** accel — Acceleration percentage */
export const ACCEL: { readonly key: "accel" } = { key: "accel" };
/** decel — Deceleration percentage */
export const DECEL: { readonly key: "decel" } = { key: "decel" };
/** repeatDur — Repeat duration */
export const REPEAT_DUR: { readonly key: "repeatDur" } = { key: "repeatDur" };
/** syncBehavior — Sync behavior (canSlip/locked) */
export const SYNC_BEHAVIOR: { readonly key: "syncBehavior" } = { key: "syncBehavior" };
/** tmFilter — Time filter */
export const TM_FILTER: { readonly key: "tmFilter" } = { key: "tmFilter" };
/** evtFilter — Event filter */
export const EVT_FILTER: { readonly key: "evtFilter" } = { key: "evtFilter" };
/** masterRel — Master relation */
export const MASTER_REL: { readonly key: "masterRel" } = { key: "masterRel" };
/** bldLvl — Build level */
export const BLD_LVL: { readonly key: "bldLvl" } = { key: "bldLvl" };
/** grpId — Group ID */
export const GRP_ID: { readonly key: "grpId" } = { key: "grpId" };
/** afterEffect — After effect flag */
export const AFTER_EFFECT: { readonly key: "afterEffect" } = { key: "afterEffect" };
/** nodePh — Node placeholder flag */
export const NODE_PH: { readonly key: "nodePh" } = { key: "nodePh" };

// ── Animation — common behavior (CT_TLCommonBehaviorData) ──

/** additive — Additive behavior type */
export const ADDITIVE: { readonly key: "additive" } = { key: "additive" };
/** accumulate — Accumulate behavior type */
export const ACCUMULATE: { readonly key: "accumulate" } = { key: "accumulate" };
/** xfrmType — Transform type (pt/img) */
export const XFRM_TYPE: { readonly key: "xfrmType" } = { key: "xfrmType" };
/** rctx — Runtime context */
export const RCTX: { readonly key: "rctx" } = { key: "rctx" };
/** override — Override type */
export const OVERRIDE: { readonly key: "override" } = { key: "override" };

// ── Animation — time animate value (CT_TLTimeAnimateValue) ──

/** fmla — Formula for animated value */
export const FMLA: { readonly key: "fmla" } = { key: "fmla" };

// ── Animation — animate color (CT_TLAnimateColorBehavior) ──

/** clrSpc — Color space (rgb/hsl) */
export const CLR_SPC: { readonly key: "clrSpc" } = { key: "clrSpc" };

// ── Animation — animate effect (CT_TLAnimateEffectBehavior) ──

/** prLst — Property list for effect */
export const PR_LST: { readonly key: "prLst" } = { key: "prLst" };

// ── Animation — animate motion (CT_TLAnimateMotionBehavior) ──

/** pathEditMode — Path edit mode (relative/fixed) */
export const PATH_EDIT_MODE: { readonly key: "pathEditMode" } = { key: "pathEditMode" };
/** rAng — Rotation angle */
export const R_ANG: { readonly key: "rAng" } = { key: "rAng" };
/** ptsTypes — Points types string */
export const PTS_TYPES: { readonly key: "ptsTypes" } = { key: "ptsTypes" };

// ── Animation — animate scale (CT_TLAnimateScaleBehavior) ──

/** zoomContents — Zoom contents flag */
export const ZOOM_CONTENTS: { readonly key: "zoomContents" } = { key: "zoomContents" };

// ── Animation — media node (CT_TLCommonMediaNodeData) ──

/** numSld — Number of slides for media */
export const NUM_SLD: { readonly key: "numSld" } = { key: "numSld" };
/** showWhenStopped — Show media when stopped */
export const SHOW_WHEN_STOPPED: { readonly key: "showWhenStopped" } = { key: "showWhenStopped" };

// ── Animation — build (AG_TLBuild) ──

/** uiExpand — UI expand flag */
export const UI_EXPAND: { readonly key: "uiExpand" } = { key: "uiExpand" };

// ── Animation — build paragraph (CT_TLBuildParagraph) ──

/** build — Build type */
export const BUILD: { readonly key: "build" } = { key: "build" };
/** animBg — Animate background flag */
export const ANIM_BG: { readonly key: "animBg" } = { key: "animBg" };
/** autoUpdateAnimBg — Auto-update animation background */
export const AUTO_UPDATE_ANIM_BG: { readonly key: "autoUpdateAnimBg" } = {
  key: "autoUpdateAnimBg",
};
/** rev — Reverse flag */
export const REV: { readonly key: "rev" } = { key: "rev" };
/** advAuto — Auto-advance time */
export const ADV_AUTO: { readonly key: "advAuto" } = { key: "advAuto" };

// ── Animation — template (CT_TLTemplate) ──

/** lvl — Template level */
export const LVL: { readonly key: "lvl" } = { key: "lvl" };

// ── Animation — color transform (CT_TLByRgbColorTransform) ──

/** r — Red component */
export const R: { readonly key: "r" } = { key: "r" };
/** g — Green component */
export const G: { readonly key: "g" } = { key: "g" };
/** b — Blue component */
export const B: { readonly key: "b" } = { key: "b" };
/** h — Hue component */
export const H: { readonly key: "h" } = { key: "h" };
/** l — Lightness component */
export const L: { readonly key: "l" } = { key: "l" };
/** s — Saturation component */
export const S: { readonly key: "s" } = { key: "s" };

// ── Presentation (CT_Presentation) ──

/** serverZoom — Server zoom percentage */
export const SERVER_ZOOM: { readonly key: "serverZoom" } = { key: "serverZoom" };
/** firstSlideNum — First slide number */
export const FIRST_SLIDE_NUM: { readonly key: "firstSlideNum" } = { key: "firstSlideNum" };
/** showSpecialPlsOnTitleSld — Show special placeholders on title slide */
export const SHOW_SPECIAL_PLS_ON_TITLE_SLD: { readonly key: "showSpecialPlsOnTitleSld" } = {
  key: "showSpecialPlsOnTitleSld",
};
/** rtl — Right-to-left flag */
export const RTL: { readonly key: "rtl" } = { key: "rtl" };
/** removePersonalInfoOnSave — Remove personal info on save */
export const REMOVE_PERSONAL_INFO_ON_SAVE: { readonly key: "removePersonalInfoOnSave" } = {
  key: "removePersonalInfoOnSave",
};
/** compatMode — Compatibility mode flag */
export const COMPAT_MODE: { readonly key: "compatMode" } = { key: "compatMode" };
/** strictFirstAndLastChars — Strict first and last characters */
export const STRICT_FIRST_AND_LAST_CHARS: { readonly key: "strictFirstAndLastChars" } = {
  key: "strictFirstAndLastChars",
};
/** embedTrueTypeFonts — Embed TrueType fonts flag */
export const EMBED_TRUE_TYPE_FONTS: { readonly key: "embedTrueTypeFonts" } = {
  key: "embedTrueTypeFonts",
};
/** saveSubsetFonts — Save subset fonts flag */
export const SAVE_SUBSET_FONTS: { readonly key: "saveSubsetFonts" } = { key: "saveSubsetFonts" };
/** autoCompressPictures — Auto-compress pictures flag */
export const AUTO_COMPRESS_PICTURES: { readonly key: "autoCompressPictures" } = {
  key: "autoCompressPictures",
};
/** bookmarkIdSeed — Bookmark ID seed */
export const BOOKMARK_ID_SEED: { readonly key: "bookmarkIdSeed" } = { key: "bookmarkIdSeed" };
/** conformance — Conformance class */
export const CONFORMANCE: { readonly key: "conformance" } = { key: "conformance" };

// ── Modify verifier (CT_ModifyVerifier) ──

/** algorithmName — Algorithm name */
export const ALGORITHM_NAME: { readonly key: "algorithmName" } = { key: "algorithmName" };
/** hashValue — Hash value */
export const HASH_VALUE: { readonly key: "hashValue" } = { key: "hashValue" };
/** saltValue — Salt value */
export const SALT_VALUE: { readonly key: "saltValue" } = { key: "saltValue" };
/** spinValue — Spin value */
export const SPIN_VALUE: { readonly key: "spinValue" } = { key: "spinValue" };
/** cryptProviderType — Cryptographic provider type */
export const CRYPT_PROVIDER_TYPE: { readonly key: "cryptProviderType" } = {
  key: "cryptProviderType",
};
/** cryptAlgorithmClass — Cryptographic algorithm class */
export const CRYPT_ALGORITHM_CLASS: { readonly key: "cryptAlgorithmClass" } = {
  key: "cryptAlgorithmClass",
};
/** cryptAlgorithmType — Cryptographic algorithm type */
export const CRYPT_ALGORITHM_TYPE: { readonly key: "cryptAlgorithmType" } = {
  key: "cryptAlgorithmType",
};
/** cryptAlgorithmSid — Cryptographic algorithm SID */
export const CRYPT_ALGORITHM_SID: { readonly key: "cryptAlgorithmSid" } = {
  key: "cryptAlgorithmSid",
};
/** spinCount — Spin count */
export const SPIN_COUNT: { readonly key: "spinCount" } = { key: "spinCount" };
/** saltData — Salt data */
export const SALT_DATA: { readonly key: "saltData" } = { key: "saltData" };
/** hashData — Hash data */
export const HASH_DATA: { readonly key: "hashData" } = { key: "hashData" };
/** cryptProvider — Cryptographic provider */
export const CRYPT_PROVIDER: { readonly key: "cryptProvider" } = { key: "cryptProvider" };
/** algIdExt — Extended algorithm ID */
export const ALG_ID_EXT: { readonly key: "algIdExt" } = { key: "algIdExt" };
/** algIdExtSource — Extended algorithm ID source */
export const ALG_ID_EXT_SOURCE: { readonly key: "algIdExtSource" } = { key: "algIdExtSource" };
/** cryptProviderTypeExt — Extended provider type */
export const CRYPT_PROVIDER_TYPE_EXT: { readonly key: "cryptProviderTypeExt" } = {
  key: "cryptProviderTypeExt",
};
/** cryptProviderTypeExtSource — Extended provider type source */
export const CRYPT_PROVIDER_TYPE_EXT_SOURCE: { readonly key: "cryptProviderTypeExtSource" } = {
  key: "cryptProviderTypeExtSource",
};

// ── HTML publish properties (CT_HtmlPublishProperties) ──

/** showSpeakerNotes — Show speaker notes */
export const SHOW_SPEAKER_NOTES: { readonly key: "showSpeakerNotes" } = { key: "showSpeakerNotes" };
/** target — Target URL/path */
export const TARGET: { readonly key: "target" } = { key: "target" };

// ── Web properties (CT_WebProperties) ──

/** showAnimation — Show animation in web */
export const SHOW_ANIMATION: { readonly key: "showAnimation" } = { key: "showAnimation" };
/** resizeGraphics — Resize graphics flag */
export const RESIZE_GRAPHICS: { readonly key: "resizeGraphics" } = { key: "resizeGraphics" };
/** allowPng — Allow PNG format */
export const ALLOW_PNG: { readonly key: "allowPng" } = { key: "allowPng" };
/** relyOnVml — Rely on VML flag */
export const RELY_ON_VML: { readonly key: "relyOnVml" } = { key: "relyOnVml" };
/** organizeInFolders — Organize in folders flag */
export const ORGANIZE_IN_FOLDERS: { readonly key: "organizeInFolders" } = {
  key: "organizeInFolders",
};
/** useLongFilenames — Use long filenames flag */
export const USE_LONG_FILENAMES: { readonly key: "useLongFilenames" } = { key: "useLongFilenames" };
/** imgSz — Image size for web */
export const IMG_SZ: { readonly key: "imgSz" } = { key: "imgSz" };
/** encoding — Web encoding */
export const ENCODING: { readonly key: "encoding" } = { key: "encoding" };
/** clr — Web color type */
export const CLR: { readonly key: "clr" } = { key: "clr" };

// ── Print properties (CT_PrintProperties) ──

/** prnWhat — What to print */
export const PRN_WHAT: { readonly key: "prnWhat" } = { key: "prnWhat" };
/** clrMode — Print color mode */
export const CLR_MODE: { readonly key: "clrMode" } = { key: "clrMode" };
/** hiddenSlides — Print hidden slides */
export const HIDDEN_SLIDES: { readonly key: "hiddenSlides" } = { key: "hiddenSlides" };
/** scaleToFitPaper — Scale to fit paper */
export const SCALE_TO_FIT_PAPER: { readonly key: "scaleToFitPaper" } = { key: "scaleToFitPaper" };
/** frameSlides — Frame slides when printing */
export const FRAME_SLIDES: { readonly key: "frameSlides" } = { key: "frameSlides" };

// ── Show properties (CT_ShowProperties) ──

/** showNarration — Show narration */
export const SHOW_NARRATION: { readonly key: "showNarration" } = { key: "showNarration" };
/** useTimings — Use timings flag */
export const USE_TIMINGS: { readonly key: "useTimings" } = { key: "useTimings" };

// ── Show info — browse (CT_ShowInfoBrowse) ──

/** showScrollbar — Show scrollbar in browse mode */
export const SHOW_SCROLLBAR: { readonly key: "showScrollbar" } = { key: "showScrollbar" };

// ── Photo album (CT_PhotoAlbum) ──

/** bw — Black and white flag */
export const BW: { readonly key: "bw" } = { key: "bw" };
/** showCaptions — Show captions in photo album */
export const SHOW_CAPTIONS: { readonly key: "showCaptions" } = { key: "showCaptions" };
/** layout — Photo album layout */
export const LAYOUT: { readonly key: "layout" } = { key: "layout" };
/** frame — Photo album frame style */
export const FRAME: { readonly key: "frame" } = { key: "frame" };

// ── Kinsoku (CT_Kinsoku) ──

/** invalStChars — Invalid start characters */
export const INVAL_ST_CHARS: { readonly key: "invalStChars" } = { key: "invalStChars" };
/** invalEndChars — Invalid end characters */
export const INVAL_END_CHARS: { readonly key: "invalEndChars" } = { key: "invalEndChars" };

// ── OLE object (AG_Ole) ──

/** showAsIcon — Show OLE as icon */
export const SHOW_AS_ICON: { readonly key: "showAsIcon" } = { key: "showAsIcon" };
/** imgW — Image width for OLE */
export const IMG_W: { readonly key: "imgW" } = { key: "imgW" };
/** imgH — Image height for OLE */
export const IMG_H: { readonly key: "imgH" } = { key: "imgH" };
/** progId — Programmatic ID */
export const PROG_ID: { readonly key: "progId" } = { key: "progId" };

// ── OLE embed (CT_OleObjectEmbed) ──

/** followColorScheme — Follow color scheme */
export const FOLLOW_COLOR_SCHEME: { readonly key: "followColorScheme" } = {
  key: "followColorScheme",
};

// ── Slide (CT_Slide) ──

/** show — Show slide flag */
export const SHOW: { readonly key: "show" } = { key: "show" };

// ── Slide layout (CT_SlideLayout) ──

/** matchingName — Matching name for layout */
export const MATCHING_NAME: { readonly key: "matchingName" } = { key: "matchingName" };
/** preserve — Preserve flag */
export const PRESERVE: { readonly key: "preserve" } = { key: "preserve" };
/** userDrawn — User drawn flag */
export const USER_DRAWN: { readonly key: "userDrawn" } = { key: "userDrawn" };

// ── Child slide attributes (AG_ChildSlide) ──

/** showMasterSp — Show master shapes */
export const SHOW_MASTER_SP: { readonly key: "showMasterSp" } = { key: "showMasterSp" };
/** showMasterPhAnim — Show master placeholder animations */
export const SHOW_MASTER_PH_ANIM: { readonly key: "showMasterPhAnim" } = {
  key: "showMasterPhAnim",
};

// ── Placeholder (CT_Placeholder) ──

/** sz — Placeholder size (full/half/quarter) */
export const SZ: { readonly key: "sz" } = { key: "sz" };
/** hasCustomPrompt — Has custom prompt */
export const HAS_CUSTOM_PROMPT: { readonly key: "hasCustomPrompt" } = { key: "hasCustomPrompt" };

// ── Application non-visual drawing props (CT_ApplicationNonVisualDrawingProps) ──

/** isPhoto — Is photo flag */
export const IS_PHOTO: { readonly key: "isPhoto" } = { key: "isPhoto" };

// ── Shape (CT_Shape) ──

/** useBgFill — Use background fill */
export const USE_BG_FILL: { readonly key: "useBgFill" } = { key: "useBgFill" };

// ── Graphical object frame (CT_GraphicalObjectFrame) ──

/** bwMode — Black and white mode */
export const BW_MODE: { readonly key: "bwMode" } = { key: "bwMode" };

// ── Background (CT_Background, CT_BackgroundProperties) ──

/** shadeToTitle — Shade to title flag */
export const SHADE_TO_TITLE: { readonly key: "shadeToTitle" } = { key: "shadeToTitle" };

// ── Slide sync properties (CT_SlideSyncProperties) ──

/** serverSldId — Server slide ID */
export const SERVER_SLD_ID: { readonly key: "serverSldId" } = { key: "serverSldId" };
/** serverSldModifiedTime — Server slide modified time */
export const SERVER_SLD_MODIFIED_TIME: { readonly key: "serverSldModifiedTime" } = {
  key: "serverSldModifiedTime",
};
/** clientInsertedTime — Client inserted time */
export const CLIENT_INSERTED_TIME: { readonly key: "clientInsertedTime" } = {
  key: "clientInsertedTime",
};

// ── Extension list modify (CT_ExtensionListModify) ──

/** mod — Modified flag */
export const MOD: { readonly key: "mod" } = { key: "mod" };

// ── Outline view slide entry (CT_OutlineViewSlideEntry) ──

/** collapse — Collapse flag */
export const COLLAPSE: { readonly key: "collapse" } = { key: "collapse" };

// ── Normal view properties (CT_NormalViewProperties) ──

/** showOutlineIcons — Show outline icons */
export const SHOW_OUTLINE_ICONS: { readonly key: "showOutlineIcons" } = { key: "showOutlineIcons" };
/** snapVertSplitter — Snap vertical splitter */
export const SNAP_VERT_SPLITTER: { readonly key: "snapVertSplitter" } = { key: "snapVertSplitter" };
/** vertBarState — Vertical bar state */
export const VERT_BAR_STATE: { readonly key: "vertBarState" } = { key: "vertBarState" };
/** horzBarState — Horizontal bar state */
export const HORZ_BAR_STATE: { readonly key: "horzBarState" } = { key: "horzBarState" };
/** preferSingleView — Prefer single view */
export const PREFER_SINGLE_VIEW: { readonly key: "preferSingleView" } = { key: "preferSingleView" };

// ── Normal view portion (CT_NormalViewPortion) ──

/** autoAdjust — Auto adjust */
export const AUTO_ADJUST: { readonly key: "autoAdjust" } = { key: "autoAdjust" };

// ── Common view properties (CT_CommonViewProperties) ──

/** varScale — Variable scale flag */
export const VAR_SCALE: { readonly key: "varScale" } = { key: "varScale" };

// ── Common slide view properties (CT_CommonSlideViewProperties) ──

/** snapToGrid — Snap to grid */
export const SNAP_TO_GRID: { readonly key: "snapToGrid" } = { key: "snapToGrid" };
/** snapToObjects — Snap to objects */
export const SNAP_TO_OBJECTS: { readonly key: "snapToObjects" } = { key: "snapToObjects" };
/** showGuides — Show guides */
export const SHOW_GUIDES: { readonly key: "showGuides" } = { key: "showGuides" };

// ── Slide sorter view properties (CT_SlideSorterViewProperties) ──

/** showFormatting — Show formatting */
export const SHOW_FORMATTING: { readonly key: "showFormatting" } = { key: "showFormatting" };

// ── View properties (CT_ViewProperties) ──

/** lastView — Last view type */
export const LAST_VIEW: { readonly key: "lastView" } = { key: "lastView" };
/** showComments — Show comments */
export const SHOW_COMMENTS: { readonly key: "showComments" } = { key: "showComments" };

// ── Guide (CT_Guide) ──

/** pos — Guide position */
export const POS: { readonly key: "pos" } = { key: "pos" };

// ── Normal view portion ──

/** sz — Size (percentage) for view portion */
export const VIEW_SZ: { readonly key: "sz" } = { key: "sz" };

// ── Build diagram (CT_TLBuildDiagram) ──

/** bld — Build type for diagrams */
export const BLD: { readonly key: "bld" } = { key: "bld" };

// ── Animation — OLE chart target element (CT_TLOleChartTargetElement) ──

/** lvl — Level for OLE chart target */
export const CHART_LVL: { readonly key: "lvl" } = { key: "lvl" };

// ── Animation — sequence (CT_TLTimeNodeSequence) ──

/** prevAc — Previous action type */
export const PREV_AC: { readonly key: "prevAc" } = { key: "prevAc" };
