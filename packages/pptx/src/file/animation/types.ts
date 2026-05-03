export type AnimationType =
    | "appear"
    | "fade"
    | "fly"
    | "wipe"
    | "dissolve"
    | "split"
    | "blinds"
    | "checker"
    | "randomBars"
    | "wheel"
    | "zoom"
    | "cover"
    | "push"
    | "strips";

export type AnimationTrigger = "onClick" | "withPrevious" | "afterPrevious";

export type AnimationDirection = "left" | "right" | "up" | "down" | "horizontal" | "vertical";

export interface IAnimationOptions {
    readonly type: AnimationType;
    readonly duration?: number;
    readonly delay?: number;
    readonly trigger?: AnimationTrigger;
    readonly direction?: AnimationDirection;
}
