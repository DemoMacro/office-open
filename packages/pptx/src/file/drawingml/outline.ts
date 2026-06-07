import { createOutline, PresetDash, LineEndType, LineEndWidth } from "@office-open/core/drawingml";
import type {
  OutlineOptions as CoreOutlineOptions,
  LineEndOptions as CoreLineEndOptions,
} from "@office-open/core/drawingml";

/**
 * PPTX-specific outline options (backward-compatible API).
 */
export interface OutlineOptions {
  width?: number;
  color?: string;
  dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

export type { CoreOutlineOptions as OutlineOptionsCore };

const DASH_STYLE_MAP: Record<string, (typeof PresetDash)[keyof typeof PresetDash]> = {
  solid: "solid",
  dash: "dash",
  dashDot: "dashDot",
  lgDash: "lgDash",
  sysDot: "sysDot",
  sysDash: "sysDash",
};

const ARROWHEAD_MAP: Record<string, (typeof LineEndType)[keyof typeof LineEndType]> = {
  triangle: "triangle",
  stealth: "stealth",
  diamond: "diamond",
  oval: "oval",
  open: "arrow",
};

function toCoreLineEnd(type: string, width?: string, length?: string): CoreLineEndOptions {
  return {
    type: ARROWHEAD_MAP[type] ?? "triangle",
    ...(width ? { width: width as (typeof LineEndWidth)[keyof typeof LineEndWidth] } : {}),
    ...(length ? { length: length as (typeof LineEndWidth)[keyof typeof LineEndWidth] } : {}),
  };
}

/**
 * Creates an outline element using pptx's simplified API.
 */
export const createOutlineCompat = (
  options: OutlineOptions = {},
  arrowheads?: {
    beginType?: string;
    endType?: string;
    width?: string;
    length?: string;
  },
) =>
  createOutline({
    width: options.width,
    ...(options.color
      ? { type: "solidFill" as const, color: { value: options.color.replace("#", "") } }
      : { type: "noFill" as const }),
    ...(options.dashStyle && {
      dash: DASH_STYLE_MAP[options.dashStyle] ?? "solid",
    }),
    ...(arrowheads?.endType
      ? { headEnd: toCoreLineEnd(arrowheads.endType, arrowheads.width, arrowheads.length) }
      : {}),
    ...(arrowheads?.beginType
      ? { tailEnd: toCoreLineEnd(arrowheads.beginType, arrowheads.width, arrowheads.length) }
      : {}),
  } satisfies CoreOutlineOptions);
