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

export type MediaAnimationType = "playAudio" | "playVideo" | "play";

export type AnimationCalcMode = "discrete" | "lin" | "fmla";

export type AnimationValueType = "str" | "num" | "clr";

export interface AnimationVariantOptions {
  bool?: boolean;
  int?: number;
  float?: number;
  string?: string;
  color?: string;
}

export interface EndConditionOptions {
  event?: string;
  delay?: string;
  timeNodeId?: number;
  runtimeNode?: "first" | "last" | "all";
}

export interface AnimationBuildOptions {
  type: "paragraph" | "diagram" | "oleChart" | "graphic";
  shapeId: number;
  groupId: number;
  uiExpand?: boolean;
  // paragraph-specific
  build?: "allAtOnce" | "p" | "cust" | "whole";
  buildLevel?: number;
  animateBackground?: boolean;
  autoUpdateAnimateBackground?: boolean;
  rev?: boolean;
  advanceAuto?: number;
  templates?: AnimationTemplateOptions[];
  // diagram-specific
  diagramBuild?: string;
  // oleChart-specific
  oleChartBuild?: string;
  oleChartAnimateBackground?: boolean;
  // graphic-specific
  graphicBuildAsOne?: boolean;
  /** Sub-element build (p:bldSub) — a:bldDgm or a:bldChart; required unless graphicBuildAsOne. */
  graphicSubBuild?: AnimationGraphicSubBuildOptions;
}

/** a:bldDgm / a:bldChart inside p:bldSub (CT_AnimationGraphicalObjectBuildProperties). */
export interface AnimationGraphicSubBuildOptions {
  /** Diagram build (a:bldDgm). */
  diagram?: {
    /** Build type (ST_AnimationDgmBuildType, default "allAtOnce") */
    build?: string;
    /** Reverse build order (`@rev`, default false) */
    reverse?: boolean;
  };
  /** Chart build (a:bldChart). */
  chart?: {
    /** Build type (ST_AnimationChartBuildType, default "allAtOnce") */
    build?: string;
    /** Animate chart background (`@animBg`, default true) */
    animateBackground?: boolean;
  };
}

export interface AnimationTemplateOptions {
  lvl?: number;
  children: AnimationOptions[];
}

export interface AnimationOptions {
  type?: AnimationType;
  duration?: number;
  delay?: number;
  trigger?: AnimationTrigger;
  direction?: AnimationDirection;
  class?: AnimationClass;
  emphasisType?: EmphasisType;
  pathType?: PathAnimationType;
  path?: string;
  /** Animation speed as integer percent (100 = normal speed). */
  speed?: number;
  repeatCount?: number;
  autoReverse?: boolean;
  color?: string;

  // Media playback animation
  mediaType?: MediaAnimationType;
  isNarration?: boolean;
  fullScreen?: boolean;
  /** Media volume as integer percent (0-100). */
  volume?: number;
  mute?: boolean;

  // Generic property animation (p:anim)
  attributeName?: string;
  calcMode?: AnimationCalcMode;
  valueType?: AnimationValueType;
  from?: string;
  to?: string;
  animBy?: string;

  // Text-level animation target
  charRange?: [number, number];
  paragraphRange?: [number, number];

  // Generic set behavior (p:set) — instant property change
  setBehavior?: {
    attributeName: string;
    toValue: string;
    toType?: "string" | "number";
  };

  // Command behavior (p:cmd) — extended command types
  commandType?: "call" | "evt" | "verb";
  command?: string;

  // Iterate container (p:iterate) — text-level iteration
  iterate?: {
    type?: "el" | "wd" | "lt";
    interval?: number;
    backwards?: boolean;
    /** Iterate step as integer percent (0–100). */
    iteratePercentage?: number;
  };

  // cTn advanced time node attributes
  /** Repeat duration ("indefinite" or milliseconds). */
  repeatDuration?: string;
  /** Acceleration as integer percent (default 0). */
  acceleration?: number;
  /** Deceleration as integer percent (default 0). */
  deceleration?: number;
  /** Restart behavior: "always", "whenNotActive", "never". */
  restart?: "always" | "whenNotActive" | "never";
  /** Sync behavior: "canSlip", "isLocked", "stoppable". */
  syncBehavior?: "canSlip" | "isLocked" | "stoppable";
  /** Time filter string. */
  timeFilter?: string;
  /** Event filter string. */
  eventFilter?: string;
  /** Display state. */
  display?: boolean;
  /** Master relationship: "clearConn", "keepConn", "resume". */
  masterRelation?: "clearConn" | "keepConn" | "resume";
  /** Build level for animation. */
  buildLevel?: number;
  /** Group ID. */
  groupId?: number;
  /** Whether this is an after-effect node. */
  afterEffect?: boolean;
  /** Node placeholder. */
  nodePlaceholder?: boolean;
  /** Automatically advance time. */
  advanceAfterTime?: string;
  /** Animate background. */
  animateBackground?: boolean;
  /** Auto-update animation background. */
  autoUpdateAnimationBackground?: boolean;

  // cBhvr behavior attributes
  /** Additive mode: "base", "sum", "repl", "mult", "none". */
  additive?: "base" | "sum" | "repl" | "mult" | "none";
  /** Accumulate mode: "none", "always". */
  accumulate?: "none" | "always";
  /** Transform type: "pt", "img". */
  transformType?: "pt" | "img";
  /** Runtime context string. */
  runtimeContext?: string;
  /** Override mode: "normal", "childStyle". */
  override?: "normal" | "childStyle";

  // Animation formula
  /** Formula for animation. */
  formula?: string;
  /** Color space for color animation. */
  colorSpace?: "rgb" | "hsl";
  /** Path edit mode. */
  pathEditMode?: "relative" | "fixed" | "none";
  /** Previous action. */
  previousAction?: "none" | "skipTimed";
  /** Points types. */
  pointsTypes?: string;
  /** Rotation angle in degrees (p:animMotion `@rAng`, ST_Angle in the XML). */
  rotationAngle?: number;
  /** Zoom contents. */
  zoomContents?: boolean;
  /** Show when stopped (for media). */
  showWhenStopped?: boolean;
  /** Number of slides. */
  numberOfSlides?: number;
  /** Property list. */
  propertyList?: string;

  // cTn time node extensions (A2)
  /** End conditions list (p:endCondLst). */
  endConditions?: EndConditionOptions[];
  /** End sync condition (p:endSync). */
  endSyncCondition?: EndConditionOptions;
  /** Sub time nodes (p:subTnLst). */
  subTimeNodes?: AnimationOptions[];
  /** Exclusive mode wrapper (p:excl). */
  exclusiveMode?: boolean;

  // Animation target extensions (A3)
  /** Ink target shape ID (p:inkTgt `@spid`). */
  inkTargetShapeId?: number;
  /** Sound target r:id (p:sndTgt). */
  soundTarget?: string;
  /** Sub-shape ID (p:subSp `@spid`). */
  subShapeId?: number;
  /** Graphic element target (p:graphicEl → a:dgm / a:chart). */
  graphicElement?: {
    /** Diagram element (a:dgm). */
    diagram?: {
      /** Diagram node GUID (a:dgm `@id`) */
      id?: string;
      /** Build step (a:dgm `@bldStep`, default "sp") */
      buildStep?: "sp" | "bg";
    };
    /** Chart element (a:chart). */
    chart?: {
      /** Series index (a:chart `@seriesIdx`, default -1) */
      seriesIndex?: number;
      /** Category index (a:chart `@categoryIdx`, default -1) */
      categoryIndex?: number;
      /** Build step (required) */
      buildStep: "category" | "ptInCategory" | "series" | "ptInSeries" | "allPts" | "gridLegend";
    };
  };
  /** OLE chart element type (p:oleChartEl `@type`). */
  oleChartElementType?: string;
  /** OLE chart element level (p:oleChartEl `@lvl`). */
  oleChartElementLevel?: number;

  // Animation variant values (A4)
  /** Variant boolean value (p:boolVal). */
  variantBool?: boolean;
  /** Variant integer value (p:intVal). */
  variantInt?: number;
  /** Variant float value (p:fltVal). */
  variantFloat?: number;
  /** Color animation from (a:srgbClr hex). */
  colorFrom?: string;
  /** Color animation to (a:srgbClr hex). */
  colorTo?: string;
  /** Color by RGB transform (p:by → p:rgb). */
  colorByRgb?: { r: string; g: string; b: string };
  /** Color by HSL transform (p:by → p:hsl). */
  colorByHsl?: { h: string; s: string; l: string };
  /** Effect progress (p:progress). */
  effectProgress?: AnimationVariantOptions;

  // Motion path extensions (A5)
  /** Motion path from point (p:from). */
  motionFrom?: { x: string; y: string };
  /** Motion path rotation center (p:rCtr). */
  motionRotationCenter?: { x: string; y: string };

  // Build animations (A1) — top-level timing container
  builds?: AnimationBuildOptions[];
}
