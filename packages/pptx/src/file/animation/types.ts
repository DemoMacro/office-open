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

export type AnimationClass = "entr" | "exit" | "emph";

export type EmphasisType =
    | "growShrink"
    | "spin"
    | "growWithTurn"
    | "colorChange"
    | "transparency"
    | "boldFlash"
    | "wave"
    | "pulse";

export type PathAnimationType =
    | "customPath"
    | "arc"
    | "bounce"
    | "circle"
    | "curve"
    | "figureEight"
    | "line"
    | "loop";

export interface IAnimationOptions {
    readonly type?: AnimationType;
    readonly duration?: number;
    readonly delay?: number;
    readonly trigger?: AnimationTrigger;
    readonly direction?: AnimationDirection;
    readonly class?: AnimationClass;
    readonly emphasisType?: EmphasisType;
    readonly pathType?: PathAnimationType;
    readonly path?: string;
    readonly speed?: number;
    readonly repeatCount?: number;
    readonly autoReverse?: boolean;
    readonly color?: string;
}
