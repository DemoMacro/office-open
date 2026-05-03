import { createOutline, PresetDash } from "@office-open/core/drawingml";
import type { OutlineOptions as CoreOutlineOptions } from "@office-open/core/drawingml";

/**
 * PPTX-specific outline options (backward-compatible API).
 */
export interface OutlineOptions {
    readonly width?: number;
    readonly color?: string;
    readonly dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

export type { CoreOutlineOptions as OutlineOptionsCore };

const DASH_STYLE_MAP: Record<string, keyof typeof PresetDash> = {
    solid: "SOLID",
    dash: "DASH",
    dashDot: "DASH_DOT",
    lgDash: "LG_DASH",
    sysDot: "SYS_DOT",
    sysDash: "SYS_DASH",
};

/**
 * Creates an outline element using pptx's simplified API.
 */
export const createOutlineCompat = (options: OutlineOptions = {}) =>
    createOutline({
        width: options.width,
        ...(options.color
            ? { type: "solidFill" as const, color: { value: options.color.replace("#", "") } }
            : { type: "noFill" as const }),
        ...(options.dashStyle && {
            dash: DASH_STYLE_MAP[options.dashStyle] ?? "SOLID",
        }),
    } satisfies CoreOutlineOptions);
