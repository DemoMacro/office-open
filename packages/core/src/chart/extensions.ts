/**
 * Chart extensions — trendlines, error bars, date axis, data labels, 3D view.
 *
 * These components are used inside chart series (trendline, errBars),
 * plot area axes (dateAx), and chart space (view3D).
 *
 * @module
 */
import { BuilderElement, XmlComponent, chartAttr, wrapEl } from "../xml-components";

// ── Trendline (CT_Trendline) ──

export const TrendlineType = {
  EXP: "exp",
  LINEAR: "linear",
  LOG: "log",
  MOVING_AVG: "movingAvg",
  POLY: "poly",
  POWER: "power",
} as const;

export type TrendlineType = (typeof TrendlineType)[keyof typeof TrendlineType];

export interface TrendlineOptions {
  /** Trendline type (default: "linear") */
  readonly type?: TrendlineType;
  /** Custom name for the trendline */
  readonly name?: string;
  /** Order for polynomial trendlines (default: 2) */
  readonly order?: number;
  /** Period for moving average trendlines (default: 2) */
  readonly period?: number;
  /** Forward forecast periods */
  readonly forward?: number;
  /** Backward forecast periods */
  readonly backward?: number;
  /** Intercept value */
  readonly intercept?: number;
  /** Display R-squared value */
  readonly dispRSqr?: boolean;
  /** Display equation */
  readonly dispEq?: boolean;
}

/**
 * c:trendline — a trendline on a chart series.
 *
 * XSD: CT_Trendline (dml-chart.xsd)
 */
export class Trendline extends XmlComponent {
  public constructor(options: TrendlineOptions) {
    super("c:trendline");

    if (options.name !== undefined) {
      this.root.push(options.name);
    }

    this.root.push(
      wrapEl("c:trendlineType", chartAttr({ val: options.type ?? TrendlineType.LINEAR })),
    );

    if (options.order !== undefined) {
      this.root.push(wrapEl("c:order", chartAttr({ val: options.order })));
    }

    if (options.period !== undefined) {
      this.root.push(wrapEl("c:period", chartAttr({ val: options.period })));
    }

    if (options.forward !== undefined) {
      this.root.push(wrapEl("c:forward", chartAttr({ val: options.forward })));
    }

    if (options.backward !== undefined) {
      this.root.push(wrapEl("c:backward", chartAttr({ val: options.backward })));
    }

    if (options.intercept !== undefined) {
      this.root.push(wrapEl("c:intercept", chartAttr({ val: options.intercept })));
    }

    if (options.dispRSqr) {
      this.root.push(wrapEl("c:dispRSqr", chartAttr({ val: 1 })));
    }

    if (options.dispEq) {
      this.root.push(wrapEl("c:dispEq", chartAttr({ val: 1 })));
    }
  }
}

// ── Error Bars (CT_ErrBars) ──

export const ErrorBarDirection = {
  BOTH: "both",
  X: "x",
  Y: "y",
} as const;

export type ErrorBarDirection = (typeof ErrorBarDirection)[keyof typeof ErrorBarDirection];

export const ErrorBarType = {
  BOTH: "both",
  MINUS: "minus",
  PLUS: "plus",
} as const;

export type ErrorBarType = (typeof ErrorBarType)[keyof typeof ErrorBarType];

export const ErrorValueType = {
  CUST: "cust",
  FIXED: "fixedVal",
  PERCENTAGE: "percentage",
  STD_DEV: "stdDev",
  STD_ERR: "stdErr",
} as const;

export type ErrorValueType = (typeof ErrorValueType)[keyof typeof ErrorValueType];

export interface ErrorBarOptions {
  /** Error bar direction (default: "y") */
  readonly direction?: ErrorBarDirection;
  /** Error bar type — both/minus/plus (required) */
  readonly barType?: ErrorBarType;
  /** Error value type (default: "stdErr") */
  readonly valueType?: ErrorValueType;
  /** Error value (for fixed/percentage/stdDev/stdErr) */
  readonly value?: number;
  /** No end cap */
  readonly noEndCap?: boolean;
}

/**
 * c:errBars — error bars on a chart series.
 *
 * XSD: CT_ErrBars (dml-chart.xsd)
 */
export class ErrorBars extends XmlComponent {
  public constructor(options: ErrorBarOptions) {
    super("c:errBars");

    if (options.direction !== undefined) {
      this.root.push(wrapEl("c:errDir", chartAttr({ val: options.direction })));
    }

    this.root.push(
      wrapEl("c:errBarType", chartAttr({ val: options.barType ?? ErrorBarType.BOTH })),
    );

    this.root.push(
      wrapEl("c:errValType", chartAttr({ val: options.valueType ?? ErrorValueType.STD_ERR })),
    );

    if (options.noEndCap) {
      this.root.push(wrapEl("c:noEndCap", chartAttr({ val: 1 })));
    }

    if (options.value !== undefined) {
      this.root.push(wrapEl("c:val", chartAttr({ val: options.value })));
    }
  }
}

// ── Data Labels (CT_DLbls / EG_DLblShared) ──

export const DataLabelPosition = {
  BEST_FIT: "bestFit",
  BOTTOM: "b",
  CENTER: "ctr",
  IN_BASE: "inBase",
  IN_END: "inEnd",
  LEFT: "l",
  OUT_END: "outEnd",
  RIGHT: "r",
  TOP: "t",
} as const;

export type DataLabelPosition = (typeof DataLabelPosition)[keyof typeof DataLabelPosition];

export interface DataLabelsOptions {
  /** Data label position */
  readonly position?: DataLabelPosition;
  /** Show value */
  readonly showVal?: boolean;
  /** Show category name */
  readonly showCatName?: boolean;
  /** Show series name */
  readonly showSerName?: boolean;
  /** Show percentage */
  readonly showPercent?: boolean;
  /** Show legend key */
  readonly showLegendKey?: boolean;
  /** Show bubble size */
  readonly showBubbleSize?: boolean;
  /** Separator character */
  readonly separator?: string;
  /** Number format */
  readonly numFmt?: string;
}

/**
 * c:dLbls — data labels settings for a chart series.
 *
 * Implements EG_DLblShared group from dml-chart.xsd.
 */
export class DataLabels extends XmlComponent {
  public constructor(options: DataLabelsOptions) {
    super("c:dLbls");

    if (options.numFmt !== undefined) {
      this.root.push(
        wrapEl("c:numFmt", chartAttr({ formatCode: options.numFmt, sourceLinked: 0 })),
      );
    }

    this.root.push(new BuilderElement({ name: "c:spPr" }));

    if (options.position !== undefined) {
      this.root.push(wrapEl("c:dLblPos", chartAttr({ val: options.position })));
    }

    if (options.showLegendKey) {
      this.root.push(wrapEl("c:showLegendKey", chartAttr({ val: 1 })));
    }
    if (options.showVal) {
      this.root.push(wrapEl("c:showVal", chartAttr({ val: 1 })));
    }
    if (options.showCatName) {
      this.root.push(wrapEl("c:showCatName", chartAttr({ val: 1 })));
    }
    if (options.showSerName) {
      this.root.push(wrapEl("c:showSerName", chartAttr({ val: 1 })));
    }
    if (options.showPercent) {
      this.root.push(wrapEl("c:showPercent", chartAttr({ val: 1 })));
    }
    if (options.showBubbleSize) {
      this.root.push(wrapEl("c:showBubbleSize", chartAttr({ val: 1 })));
    }
    if (options.separator !== undefined) {
      this.root.push(new SeparatorElement(options.separator));
    }
  }
}

class SeparatorElement extends XmlComponent {
  public constructor(val: string) {
    super("c:separator");
    this.root.push(val);
  }
}

// ── Date Axis (CT_DateAx) ──

export const TimeUnit = {
  DAYS: "days",
  MONTHS: "months",
  YEARS: "years",
} as const;

export type TimeUnit = (typeof TimeUnit)[keyof typeof TimeUnit];

/**
 * c:dateAx — date axis for time-series data.
 *
 * XSD: CT_DateAx (dml-chart.xsd)
 */
export class DateAx extends XmlComponent {
  public constructor(axId: number, crossAx: number) {
    super("c:dateAx");
    // EG_AxShared
    this.root.push(wrapEl("c:axId", chartAttr({ val: axId })));
    this.root.push(new Scaling());
    this.root.push(wrapEl("c:delete", chartAttr({ val: 0 })));
    this.root.push(wrapEl("c:axPos", chartAttr({ val: "b" })));
    this.root.push(wrapEl("c:numFmt", chartAttr({ formatCode: "General", sourceLinked: 1 })));
    this.root.push(new BuilderElement({ name: "c:spPr" }));
    this.root.push(wrapEl("c:crossAx", chartAttr({ val: crossAx })));
    this.root.push(wrapEl("c:crosses", chartAttr({ val: "autoZero" })));
    // DateAx-specific
    this.root.push(wrapEl("c:auto", chartAttr({ val: 1 })));
    this.root.push(wrapEl("c:lblOffset", chartAttr({ val: 100 })));
  }
}

class Scaling extends XmlComponent {
  public constructor() {
    super("c:scaling");
    this.root.push(wrapEl("c:orientation", chartAttr({ val: "minMax" })));
  }
}

// ── 3D View (CT_View3D) ──

export interface View3DOptions {
  /** X rotation in degrees (-90 to 90, default: 30) */
  readonly rotX?: number;
  /** Y rotation in degrees (0 to 360, default: 0) */
  readonly rotY?: number;
  /** Height percentage (5 to 500) */
  readonly heightPercent?: number;
  /** Depth percentage (20 to 2000, default: 100) */
  readonly depthPercent?: number;
  /** Right angle axes */
  readonly rightAngleAxes?: boolean;
  /** Perspective (0 to 240) */
  readonly perspective?: number;
}

/**
 * c:view3D — 3D view properties for 3D charts.
 *
 * XSD: CT_View3D (dml-chart.xsd)
 */
export class View3D extends XmlComponent {
  public constructor(options?: View3DOptions) {
    super("c:view3D");
    const opts = options ?? {};
    if (opts.rotX !== undefined) {
      this.root.push(wrapEl("c:rotX", chartAttr({ val: opts.rotX })));
    }
    if (opts.heightPercent !== undefined) {
      this.root.push(wrapEl("c:hPercent", chartAttr({ val: opts.heightPercent })));
    }
    if (opts.rotY !== undefined) {
      this.root.push(wrapEl("c:rotY", chartAttr({ val: opts.rotY })));
    }
    if (opts.depthPercent !== undefined) {
      this.root.push(wrapEl("c:depthPercent", chartAttr({ val: opts.depthPercent })));
    }
    if (opts.rightAngleAxes !== undefined) {
      this.root.push(wrapEl("c:rAngAx", chartAttr({ val: opts.rightAngleAxes ? 1 : 0 })));
    }
    if (opts.perspective !== undefined) {
      this.root.push(wrapEl("c:perspective", chartAttr({ val: opts.perspective })));
    }
  }
}
