import type { Context } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";

const ORIENT_TYPES = new Set(["blinds", "checker", "comb", "randomBar"]);

const SIDE_DIR_TYPES = new Set(["push", "wipe"]);

const EIGHT_DIR_TYPES = new Set(["cover", "pull"]);

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

const DIRECTION_MAP: Record<TransitionDirection, string> = {
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
  readonly type?: TransitionType;
  readonly speed?: "slow" | "med" | "fast";
  readonly advanceOnClick?: boolean;
  readonly advanceAfterTime?: number;
  readonly direction?: TransitionDirection;
  readonly orient?: "horz" | "vert";
  readonly thruBlk?: boolean;
  readonly spokes?: number;
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

  if (ORIENT_TYPES.has(type) && dir && dir !== "horz") {
    attrs.push(`dir="${dir}"`);
  } else if (SIDE_DIR_TYPES.has(type) && dir && dir !== "l") {
    attrs.push(`dir="${dir}"`);
  } else if (EIGHT_DIR_TYPES.has(type) && dir && dir !== "l") {
    attrs.push(`dir="${dir}"`);
  } else if (type === "strips" && dir && dir !== "lu") {
    attrs.push(`dir="${dir}"`);
  } else if ((type === "fade" || type === "cut") && thruBlk !== undefined) {
    attrs.push(`thruBlk="${thruBlk ? 1 : 0}"`);
  } else if (type === "split") {
    if (orient && orient !== "horz") attrs.push(`orient="${orient}"`);
    if (dir && dir !== "out") attrs.push(`dir="${dir}"`);
  } else if (type === "wheel") {
    if (spokes !== undefined && spokes !== 4) attrs.push(`spokes="${spokes}"`);
  } else if (type === "zoom" && dir && dir !== "in") {
    attrs.push(`dir="${dir}"`);
  }

  return attrs.length > 0 ? `<p:${type} ${attrs.join(" ")}/>` : `<p:${type}/>`;
}

export function buildTransition(options: TransitionOptions): string {
  const attrParts: string[] = [];
  if (options.speed) attrParts.push(`spd="${options.speed}"`);
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

  if (attrParts.length === 0 && children.length === 0) {
    return "<p:transition/>";
  }

  const attrStr = attrParts.length > 0 ? ` ${attrParts.join(" ")}` : "";
  if (children.length === 0) {
    return `<p:transition${attrStr}/>`;
  }
  return `<p:transition${attrStr}>${children.join("")}</p:transition>`;
}

/**
 * p:transition — Slide transition effect.
 * Lazy: stores options, builds XML in toXml.
 */
export class Transition extends BaseXmlComponent {
  private readonly options: TransitionOptions;

  public constructor(options: TransitionOptions = {}) {
    super("p:transition");
    this.options = options;
  }

  public override toXml(_context: Context): string {
    return buildTransition(this.options);
  }
}
