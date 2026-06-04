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

export const ADV_CLICK: { readonly key: "advClick" } = { key: "advClick" };
export const ADV_TM: { readonly key: "advTm" } = { key: "advTm" };
export const DIR: { readonly key: "dir" } = { key: "dir" };
export const ORIENT: { readonly key: "orient" } = { key: "orient" };
export const THRU_BLK: { readonly key: "thruBlk" } = { key: "thruBlk" };
export const SPOKES: { readonly key: "spokes" } = { key: "spokes" };
export const LOOP: { readonly key: "loop" } = { key: "loop" };

// ── Animation — common time node (CT_TLCommonTimeNodeData) ──

export const ACCEL: { readonly key: "accel" } = { key: "accel" };
export const DECEL: { readonly key: "decel" } = { key: "decel" };
export const REPEAT_DUR: { readonly key: "repeatDur" } = { key: "repeatDur" };
export const SYNC_BEHAVIOR: { readonly key: "syncBehavior" } = { key: "syncBehavior" };
export const TM_FILTER: { readonly key: "tmFilter" } = { key: "tmFilter" };
export const EVT_FILTER: { readonly key: "evtFilter" } = { key: "evtFilter" };
export const MASTER_REL: { readonly key: "masterRel" } = { key: "masterRel" };
export const BLD_LVL: { readonly key: "bldLvl" } = { key: "bldLvl" };
export const GRP_ID: { readonly key: "grpId" } = { key: "grpId" };
export const AFTER_EFFECT: { readonly key: "afterEffect" } = { key: "afterEffect" };
export const NODE_PH: { readonly key: "nodePh" } = { key: "nodePh" };

// ── Animation — common behavior (CT_TLCommonBehaviorData) ──

export const ADDITIVE: { readonly key: "additive" } = { key: "additive" };
export const ACCUMULATE: { readonly key: "accumulate" } = { key: "accumulate" };
export const XFRM_TYPE: { readonly key: "xfrmType" } = { key: "xfrmType" };
export const RCTX: { readonly key: "rctx" } = { key: "rctx" };
export const OVERRIDE: { readonly key: "override" } = { key: "override" };

// ── Animation — time animate value ──

export const FMLA: { readonly key: "fmla" } = { key: "fmla" };

// ── Animation — animate color ──

export const CLR_SPC: { readonly key: "clrSpc" } = { key: "clrSpc" };

// ── Animation — animate effect ──

export const PR_LST: { readonly key: "prLst" } = { key: "prLst" };

// ── Animation — animate motion ──

export const PATH_EDIT_MODE: { readonly key: "pathEditMode" } = { key: "pathEditMode" };
export const R_ANG: { readonly key: "rAng" } = { key: "rAng" };
export const PTS_TYPES: { readonly key: "ptsTypes" } = { key: "ptsTypes" };

// ── Animation — animate scale ──

export const ZOOM_CONTENTS: { readonly key: "zoomContents" } = { key: "zoomContents" };

// ── Animation — media node ──

export const NUM_SLD: { readonly key: "numSld" } = { key: "numSld" };
export const SHOW_WHEN_STOPPED: { readonly key: "showWhenStopped" } = { key: "showWhenStopped" };

// ── Animation — build ──

export const UI_EXPAND: { readonly key: "uiExpand" } = { key: "uiExpand" };
export const BUILD: { readonly key: "build" } = { key: "build" };
export const ANIM_BG: { readonly key: "animBg" } = { key: "animBg" };
export const AUTO_UPDATE_ANIM_BG: { readonly key: "autoUpdateAnimBg" } = {
  key: "autoUpdateAnimBg",
};
export const REV: { readonly key: "rev" } = { key: "rev" };
export const ADV_AUTO: { readonly key: "advAuto" } = { key: "advAuto" };

// ── Animation — template ──

export const LVL: { readonly key: "lvl" } = { key: "lvl" };

// ── Animation — color transform ──

export const R: { readonly key: "r" } = { key: "r" };
export const G: { readonly key: "g" } = { key: "g" };
export const B: { readonly key: "b" } = { key: "b" };
export const H: { readonly key: "h" } = { key: "h" };
export const L: { readonly key: "l" } = { key: "l" };
export const S: { readonly key: "s" } = { key: "s" };

// ── Kinsoku ──

export const INVAL_ST_CHARS: { readonly key: "invalStChars" } = { key: "invalStChars" };
export const INVAL_END_CHARS: { readonly key: "invalEndChars" } = { key: "invalEndChars" };

// ── OLE object ──

export const SHOW_AS_ICON: { readonly key: "showAsIcon" } = { key: "showAsIcon" };
export const IMG_W: { readonly key: "imgW" } = { key: "imgW" };
export const IMG_H: { readonly key: "imgH" } = { key: "imgH" };
export const PROG_ID: { readonly key: "progId" } = { key: "progId" };
export const FOLLOW_COLOR_SCHEME: { readonly key: "followColorScheme" } = {
  key: "followColorScheme",
};

// ── Slide ──

export const SHOW: { readonly key: "show" } = { key: "show" };

// ── Slide layout ──

export const MATCHING_NAME: { readonly key: "matchingName" } = { key: "matchingName" };
export const PRESERVE: { readonly key: "preserve" } = { key: "preserve" };
export const USER_DRAWN: { readonly key: "userDrawn" } = { key: "userDrawn" };

// ── Child slide attributes ──

export const SHOW_MASTER_SP: { readonly key: "showMasterSp" } = { key: "showMasterSp" };
export const SHOW_MASTER_PH_ANIM: { readonly key: "showMasterPhAnim" } = {
  key: "showMasterPhAnim",
};

// ── Placeholder ──

export const SZ: { readonly key: "sz" } = { key: "sz" };
export const HAS_CUSTOM_PROMPT: { readonly key: "hasCustomPrompt" } = { key: "hasCustomPrompt" };

// ── Application non-visual drawing props ──

export const IS_PHOTO: { readonly key: "isPhoto" } = { key: "isPhoto" };

// ── Shape ──

export const USE_BG_FILL: { readonly key: "useBgFill" } = { key: "useBgFill" };

// ── Graphical object frame ──

export const BW_MODE: { readonly key: "bwMode" } = { key: "bwMode" };

// ── Background ──

export const SHADE_TO_TITLE: { readonly key: "shadeToTitle" } = { key: "shadeToTitle" };

// ── Slide sync properties ──

export const SERVER_SLD_ID: { readonly key: "serverSldId" } = { key: "serverSldId" };
export const SERVER_SLD_MODIFIED_TIME: { readonly key: "serverSldModifiedTime" } = {
  key: "serverSldModifiedTime",
};
export const CLIENT_INSERTED_TIME: { readonly key: "clientInsertedTime" } = {
  key: "clientInsertedTime",
};

// ── Extension list modify ──

export const MOD: { readonly key: "mod" } = { key: "mod" };

// ── Outline view slide entry ──

export const COLLAPSE: { readonly key: "collapse" } = { key: "collapse" };

// ── Slide sorter view properties ──

export const SHOW_FORMATTING: { readonly key: "showFormatting" } = { key: "showFormatting" };

// ── Guide ──

export const POS: { readonly key: "pos" } = { key: "pos" };

// ── Build diagram ──

export const BLD: { readonly key: "bld" } = { key: "bld" };

// ── Animation — OLE chart target element ──

export const CHART_LVL: { readonly key: "lvl" } = { key: "lvl" };

// ── Animation — sequence ──

export const PREV_AC: { readonly key: "prevAc" } = { key: "prevAc" };
