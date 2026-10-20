import type { DataType } from "@office-open/core";
import { xsdOrient } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import type { WriteContext } from "@office-open/core/descriptor";

const ORIENT_TYPES = new Set(["blinds", "checker", "comb", "randomBar"]);

const SIDE_DIR_TYPES = new Set(["push", "wipe"]);

const EIGHT_DIR_TYPES = new Set(["cover", "pull"]);

/** Slide-change effect (ST_Transition): "newsflash" spinning blur, "plus" plus-shaped reveal, "wedge" wedge-shaped wipe, "comb" comb-tooth reveal, "randomBar" random bars. */
export type TransitionType =
  | "fade"
  | "push"
  | "wipe"
  | "split"
  | "blinds"
  | "checker"
  | "comb"
  | "randomBar"
  | "cover"
  | "pull"
  | "strips"
  | "wheel"
  | "zoom"
  | "circle"
  | "dissolve"
  | "diamond"
  | "newsflash"
  | "plus"
  | "wedge"
  | "random"
  | "cut";

/** Slide-transition motion direction; diagonal corners like "leftUp" apply to cover/pull, "in"/"out" to circle/plus/wedge/diamond. */
export type TransitionDirection =
  | "left"
  | "up"
  | "right"
  | "down"
  | "leftUp"
  | "rightUp"
  | "leftDown"
  | "rightDown"
  | "out"
  | "in";

export const DIRECTION_MAP: Record<TransitionDirection, string> = {
  left: "l",
  up: "u",
  right: "r",
  down: "d",
  leftUp: "lu",
  rightUp: "ru",
  leftDown: "ld",
  rightDown: "rd",
  out: "out",
  in: "in",
};

export interface TransitionOptions {
  /** Slide-change effect; see TransitionType. */
  type?: TransitionType;
  /** Transition playback speed. */
  speed?: "slow" | "medium" | "fast";
  /** Transition duration in milliseconds (p14:dur). */
  duration?: number;
  advanceOnClick?: boolean;
  advanceAfterTime?: number;
  /** Motion direction; see TransitionDirection. */
  direction?: TransitionDirection;
  /** Split/wipe orientation. */
  orient?: "horizontal" | "vertical";
  thruBlk?: boolean;
  spokes?: number;
  startSound?: {
    /** Audio content: base64 string, data URL, or raw bytes. */
    data: DataType;
    /** Audio container format — becomes the media part extension. */
    type: "mp3" | "wav" | "wma" | "aac";
    /** Source media file name, pinned so relationship targets stay stable. */
    fileName?: string;
    name?: string;
    loop?: boolean;
  };
  stopPreviousSound?: boolean;
}

function buildTransitionElement(
  type: TransitionType,
  direction?: TransitionDirection,
  orient?: string,
  thruBlk?: boolean,
  spokes?: number,
): string {
  const dir = direction ? DIRECTION_MAP[direction] : undefined;
  const attrs: string[] = [];

  if (ORIENT_TYPES.has(type)) {
    if (orient !== undefined) attrs.push(`dir="${xsdOrient.to(orient)}"`);
  } else if (
    ((SIDE_DIR_TYPES.has(type) || EIGHT_DIR_TYPES.has(type)) && dir !== undefined) ||
    (type === "strips" && dir !== undefined)
  ) {
    attrs.push(`dir="${dir}"`);
  } else if ((type === "fade" || type === "cut") && thruBlk !== undefined) {
    attrs.push(`thruBlk="${thruBlk ? 1 : 0}"`);
  } else if (type === "split") {
    if (orient !== undefined) attrs.push(`orient="${xsdOrient.to(orient)}"`);
    if (dir !== undefined) attrs.push(`dir="${dir}"`);
  } else if (type === "wheel") {
    if (spokes !== undefined) attrs.push(`spokes="${spokes}"`);
  } else if (type === "zoom" && dir !== undefined) {
    attrs.push(`dir="${dir}"`);
  }

  return attrs.length > 0 ? `<p:${type} ${attrs.join(" ")}/>` : `<p:${type}/>`;
}

export function buildTransition(options: TransitionOptions, ctx?: WriteContext): string {
  const attrParts: string[] = [];
  if (options.speed)
    // OOXML ST_TransitionSpeed uses "med"; the friendly API token is "medium".
    attrParts.push(`spd="${options.speed === "medium" ? "med" : options.speed}"`);
  if (options.duration !== undefined) {
    attrParts.push(
      `xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="${options.duration}"`,
    );
  }
  if (options.advanceOnClick !== undefined)
    attrParts.push(`advClick="${options.advanceOnClick ? 1 : 0}"`);
  if (options.advanceAfterTime !== undefined) attrParts.push(`advTm="${options.advanceAfterTime}"`);

  const children: string[] = [];
  if (options.type) {
    children.push(
      buildTransitionElement(
        options.type,
        options.direction,
        options.orient,
        options.thruBlk,
        options.spokes,
      ),
    );
  }

  // Sound action (sndAc: stSnd | endSnd). The sound bytes are registered as
  // media through the write context; the {audio:…} placeholder is rewritten to
  // a real relationship id (and the audio relationship added) by the compiler.
  if (options.startSound && ctx) {
    const ref = ctx.addMedia(
      toUint8Array(options.startSound.data),
      options.startSound.type,
      options.startSound.fileName,
    );
    const audioRef = `{audio:${ref.slice(1, -1)}}`;
    const sndAttrs: string[] = [`r:embed="${audioRef}"`];
    if (options.startSound.name) sndAttrs.push(` name="${options.startSound.name}"`);
    const stSndAttrs: string[] = [];
    if (options.startSound.loop) stSndAttrs.push(' loop="1"');
    children.push(
      `<p:sndAc><p:stSnd${stSndAttrs.join("")}>` +
        `<p:snd ${sndAttrs.join(" ")}/>` +
        `</p:stSnd></p:sndAc>`,
    );
  } else if (options.stopPreviousSound) {
    children.push("<p:sndAc><p:endSnd/></p:sndAc>");
  }

  if (attrParts.length === 0 && children.length === 0) {
    return "<p:transition/>";
  }

  const attrStr = attrParts.length > 0 ? ` ${attrParts.join(" ")}` : "";
  if (children.length === 0) {
    return `<p:transition${attrStr}/>`;
  }
  return `<p:transition${attrStr}>${children.join("")}</p:transition>`;
}
