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

export interface AnimationOptions {
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

  // Generic set behavior (p:set) — instant property change
  readonly setBehavior?: {
    readonly attributeName: string;
    readonly toValue: string;
    readonly toType?: "string" | "number";
  };

  // Command behavior (p:cmd) — extended command types
  readonly commandType?: "call" | "evt" | "verb";
  readonly command?: string;

  // Iterate container (p:iterate) — text-level iteration
  readonly iterate?: {
    readonly type?: "el" | "wd" | "lt";
    readonly interval?: number;
    readonly backwards?: boolean;
  };

  // cTn advanced time node attributes
  /** Repeat duration ("indefinite" or milliseconds). */
  readonly repeatDuration?: string;
  /** Acceleration percentage (0–100000, default 0). */
  readonly acceleration?: number;
  /** Deceleration percentage (0–100000, default 0). */
  readonly deceleration?: number;
  /** Restart behavior: "always", "whenNotActive", "never". */
  readonly restart?: "always" | "whenNotActive" | "never";
  /** Sync behavior: "canSlip", "isLocked", "stoppable". */
  readonly syncBehavior?: "canSlip" | "isLocked" | "stoppable";
  /** Time filter string. */
  readonly timeFilter?: string;
  /** Event filter string. */
  readonly eventFilter?: string;
  /** Display state. */
  readonly display?: boolean;
  /** Master relationship: "clearConn", "keepConn", "resume". */
  readonly masterRelation?: "clearConn" | "keepConn" | "resume";
  /** Build level for animation. */
  readonly buildLevel?: number;
  /** Group ID. */
  readonly groupId?: number;
  /** Whether this is an after-effect node. */
  readonly afterEffect?: boolean;
  /** Node placeholder. */
  readonly nodePlaceholder?: boolean;
  /** Automatically advance time. */
  readonly advanceAfterTime?: string;
  /** Animate background. */
  readonly animateBackground?: boolean;
  /** Auto-update animation background. */
  readonly autoUpdateAnimationBackground?: boolean;

  // cBhvr behavior attributes
  /** Additive mode: "base", "sum", "repl", "mult", "none". */
  readonly additive?: "base" | "sum" | "repl" | "mult" | "none";
  /** Accumulate mode: "none", "always". */
  readonly accumulate?: "none" | "always";
  /** Transform type: "pt", "img". */
  readonly transformType?: "pt" | "img";
  /** Runtime context string. */
  readonly runtimeContext?: string;
  /** Override mode: "normal", "childStyle". */
  readonly override?: "normal" | "childStyle";

  // Animation build and formula
  /** Paragraph build type: "whole", "allAtOnce", "p", "cust". */
  readonly paragraphBuild?: "whole" | "allAtOnce" | "p" | "cust";
  /** Formula for animation. */
  readonly formula?: string;
  /** Color space for color animation. */
  readonly colorSpace?: "rgb" | "hsl";
  /** Path edit mode. */
  readonly pathEditMode?: "relative" | "fixed" | "none";
  /** Previous action. */
  readonly previousAction?: "none" | "skipTimed";
  /** Points types. */
  readonly pointsTypes?: string;
  /** Rotation angle. */
  readonly rotationAngle?: number;
  /** Zoom contents. */
  readonly zoomContents?: boolean;
  /** Show when stopped (for media). */
  readonly showWhenStopped?: boolean;
  /** Number of slides. */
  readonly numberOfSlides?: number;
  /** Property list. */
  readonly propertyList?: string;
}
