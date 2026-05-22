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

export type AnimationClass = "entr" | "exit" | "emph" | "mediacall";

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

export type MediaAnimationType = "playAudio" | "playVideo";

export type AnimationCalcMode = "discrete" | "lin" | "fmla";

export type AnimationValueType = "str" | "num" | "clr";

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

  // Media playback animation
  readonly mediaType?: MediaAnimationType;
  readonly isNarration?: boolean;
  readonly fullScreen?: boolean;
  readonly volume?: number;
  readonly mute?: boolean;

  // Generic property animation (p:anim)
  readonly attributeName?: string;
  readonly calcMode?: AnimationCalcMode;
  readonly valueType?: AnimationValueType;
  readonly from?: string;
  readonly to?: string;
  readonly animBy?: string;

  // Text-level animation target
  readonly charRange?: readonly [number, number];
  readonly paragraphRange?: readonly [number, number];
}
