/**
 * Bidirectional ChartEx codec.
 *
 * @module
 */

import {
  attr,
  children as rawChildren,
  element as rawElement,
  findChild as rawFindChild,
  stringifyElement,
  textOf,
} from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../descriptor";
import { fillDesc } from "../drawing/fill/fill-descriptors";
import type { FillOptions } from "../drawing/fill/fill-options";
import { shapePropertiesDesc } from "../drawing/shape-properties-desc";
import type { ShapePropertiesOptions } from "../drawing/shape-properties-desc";
import { textBodyDesc } from "../drawing/text/text-body";
import type { TextBodyOptions } from "../drawing/text/text-body";
import { parseColorMapping, stringifyColorMapping } from "../theme/color-mapping";
import { ChartExParseError } from "./chart-ex";
import type {
  ChartExAlignment,
  ChartExAutomaticNumber,
  ChartExAxis,
  ChartExAxisUnit,
  ChartExBinning,
  ChartExBinding,
  ChartExChart,
  ChartExChartData,
  ChartExDashboard,
  ChartExData,
  ChartExDataLabel,
  ChartExDataLabels,
  ChartExDimension,
  ChartExExtension,
  ChartExFormatOverride,
  ChartExFormula,
  ChartExGeography,
  ChartExGeoCache,
  ChartExGridlines,
  ChartExLabelPosition,
  ChartExLabelVisibility,
  ChartExNumberFormat,
  ChartExNumericValue,
  ChartExPageSetup,
  ChartExPrintSettings,
  ChartExQuartileMethod,
  ChartExSeries,
  ChartExSeriesLayoutProperties,
  ChartExSide,
  ChartExSpaceOptions,
  ChartExStringValue,
  ChartExText,
  ChartExTickMarks,
  ChartExTickMarkType,
  ChartExValueColors,
} from "./chart-ex";

const SIDES: Record<ChartExSide, string> = { left: "l", top: "t", right: "r", bottom: "b" };
const SIDES_FROM: Record<string, ChartExSide> = { l: "left", t: "top", r: "right", b: "bottom" };
const ALIGNMENTS: Record<ChartExAlignment, string> = {
  minimum: "min",
  center: "ctr",
  maximum: "max",
};
const ALIGNMENTS_FROM: Record<string, ChartExAlignment> = {
  min: "minimum",
  ctr: "center",
  max: "maximum",
};
const DIRECTIONS = { column: "col", row: "row" } as const;
const TICK_MARKS: Record<ChartExTickMarkType, string> = {
  inside: "in",
  outside: "out",
  cross: "cross",
  none: "none",
};
const TICK_MARKS_FROM: Record<string, ChartExTickMarkType> = {
  in: "inside",
  out: "outside",
  cross: "cross",
  none: "none",
};
const LABEL_POSITIONS: Record<ChartExLabelPosition, string> = {
  bestFit: "bestFit",
  bottom: "b",
  center: "ctr",
  insideBase: "inBase",
  insideEnd: "inEnd",
  left: "l",
  outsideEnd: "outEnd",
  right: "r",
  top: "t",
};
const LABEL_POSITIONS_FROM: Record<string, ChartExLabelPosition> = {
  bestFit: "bestFit",
  b: "bottom",
  ctr: "center",
  inBase: "insideBase",
  inEnd: "insideEnd",
  l: "left",
  outEnd: "outsideEnd",
  r: "right",
  t: "top",
};
const NUMERIC_DIMENSIONS = new Set(["val", "x", "xVal", "y", "yVal", "size", "degree", "colorVal"]);

const element = (
  name: string,
  attrs?: Readonly<Record<string, string | number | boolean | undefined>>,
  children?: readonly string[],
): string => rawElement(name.includes(":") ? name : `cx:${name}`, attrs, children);

const findChild = (parent: XmlElement | undefined, name: string): XmlElement | undefined =>
  rawFindChild(parent, name.includes(":") ? name : `cx:${name}`);

const children = (parent: XmlElement | undefined, name: string): XmlElement[] =>
  rawChildren(parent, name.includes(":") ? name : `cx:${name}`);

function isName(el: XmlElement, name: string): boolean {
  return el.name === name || el.name === `cx:${name}`;
}

type ChartExNumericLevelLike = {
  pointCount: number;
  name?: string;
  formatCode?: string;
  points?: ChartExNumericValue[];
};
type ChartExStringLevelLike = {
  pointCount: number;
  name?: string;
  formatCode?: string;
  points?: ChartExStringValue[];
};
type ChartExGridlinesLike = ChartExGridlines;
type ChartExTickMarksLike = ChartExTickMarks;
type ChartExFill = FillOptions;

function required(el: XmlElement | undefined, name: string, path: string): string {
  const value = attr(el, name) ?? attr(el, `cx:${name}`);
  if (value === undefined) throw new Error(`ChartEx ${path} is missing @${name}`);
  return String(value);
}

function number(value: string | number | undefined, path: string, name: string): number {
  const parsed = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(parsed)) throw new Error(`ChartEx ${path} has invalid @${name}`);
  return parsed;
}

function requiredNumber(el: XmlElement, name: string, path: string): number {
  return number(required(el, name, path), path, name);
}

function boolean(el: XmlElement, name: string, path: string): boolean | undefined {
  const value = attr(el, name) ?? attr(el, `cx:${name}`);
  if (value === undefined) return undefined;
  const normalized = String(value).toLowerCase();
  if (normalized === "true" || normalized === "1") return true;
  if (normalized === "false" || normalized === "0") return false;
  throw new Error(`ChartEx ${path} has invalid boolean @${name}`);
}

function extensionsXml(list: ChartExExtension[] | undefined): string {
  if (!list?.length) return "";
  return element(
    "extLst",
    undefined,
    list.map((item) => element("ext", { uri: item.uri }, [item.content ?? ""])),
  );
}

function parseExtensions(el: XmlElement): ChartExExtension[] | undefined {
  const list = el.elements?.filter((child) => isName(child, "ext")) ?? [];
  if (!list.length) return undefined;
  return list.map((item) => ({
    ...(item.attributes?.uri === undefined ? {} : { uri: String(item.attributes.uri) }),
    ...(item.elements?.length ? { content: item.elements.map(stringifyElement).join("") } : {}),
  }));
}

function optionalExtensions(el: XmlElement): { extensions?: ChartExExtension[] } {
  const extensions = parseExtensions(el);
  return extensions ? { extensions } : {};
}

function formulaXml(tag: string, formula: ChartExFormula | undefined): string {
  if (!formula) return "";
  return element(
    tag,
    { dir: formula.direction === undefined ? undefined : DIRECTIONS[formula.direction] },
    [formula.text],
  );
}

function parseFormula(el: XmlElement | undefined): ChartExFormula | undefined {
  if (!el) return undefined;
  const direction = attr(el, "dir");
  return {
    text: textOf(el),
    ...(direction === undefined ? {} : { direction: direction === "row" ? "row" : "column" }),
  };
}

function shapeXml(options: ShapePropertiesOptions | undefined): string {
  return options
    ? element("spPr", undefined, [shapePropertiesDesc.stringify(options, {} as WriteContext) ?? ""])
    : "";
}

function parseShape(el: XmlElement | undefined): ShapePropertiesOptions | undefined {
  return el ? shapePropertiesDesc.parse(el, {} as ReadContext) : undefined;
}

function textPropertiesXml(options: TextBodyOptions | undefined): string {
  return options
    ? element("txPr", undefined, [textBodyDesc.stringify(options, {} as WriteContext) ?? ""])
    : "";
}

function parseTextProperties(el: XmlElement | undefined): TextBodyOptions | undefined {
  return el ? textBodyDesc.parse(el, {} as ReadContext) : undefined;
}

function numberFormatXml(options: ChartExNumberFormat | undefined): string {
  if (!options) return "";
  return element("numFmt", { formatCode: options.formatCode, sourceLinked: options.sourceLinked });
}

function parseNumberFormat(
  el: XmlElement | undefined,
  path: string,
): ChartExNumberFormat | undefined {
  if (!el) return undefined;
  const sourceLinked = boolean(el, "sourceLinked", path);
  return {
    formatCode: required(el, "formatCode", path),
    ...(sourceLinked === undefined ? {} : { sourceLinked }),
  };
}

function textXml(options: ChartExText | undefined): string {
  if (!options) return "";
  if ("rich" in options)
    return element("tx", undefined, [
      element("rich", undefined, [textBodyDesc.stringify(options.rich, {} as WriteContext) ?? ""]),
    ]);
  return element("tx", undefined, [
    element("txData", undefined, [
      formulaXml("f", options.formula),
      options.value === undefined ? "" : element("v", undefined, [options.value]),
    ]),
  ]);
}

function parseText(el: XmlElement | undefined): ChartExText | undefined {
  if (!el) return undefined;
  const data = findChild(el, "txData");
  if (data) {
    const formula = parseFormula(findChild(data, "f"));
    const value = findChild(data, "v");
    return {
      ...(formula ? { formula } : {}),
      ...(value ? { value: textOf(value) } : {}),
    };
  }
  const rich = findChild(el, "rich");
  if (rich) return { rich: textBodyDesc.parse(rich, {} as ReadContext) };
  throw new Error("ChartEx tx is missing txData or rich");
}

function containerXml(options: {
  text?: ChartExText;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  offset?: { top: number; left: number };
  extensions?: ChartExExtension[];
}): string {
  return [
    textXml(options.text),
    shapeXml(options.shapeProperties),
    textPropertiesXml(options.textProperties),
    options.offset ? element("offset", { top: options.offset.top, left: options.offset.left }) : "",
    extensionsXml(options.extensions),
  ].join("");
}

function parseContainer(el: XmlElement): {
  text?: ChartExText;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  offset?: { top: number; left: number };
  extensions?: ChartExExtension[];
} {
  const text = parseText(findChild(el, "tx"));
  const shapeProperties = parseShape(findChild(el, "spPr"));
  const textProperties = parseTextProperties(findChild(el, "txPr"));
  const offsetElement = findChild(el, "offset");
  const extensions = parseExtensions(el);
  return {
    ...(text ? { text } : {}),
    ...(shapeProperties ? { shapeProperties } : {}),
    ...(textProperties ? { textProperties } : {}),
    ...(offsetElement
      ? {
          offset: {
            top: requiredNumber(offsetElement, "top", "offset"),
            left: requiredNumber(offsetElement, "left", "offset"),
          },
        }
      : {}),
    ...(extensions ? { extensions } : {}),
  };
}

function positionXml(options: {
  position?: ChartExSide;
  alignment?: ChartExAlignment;
  overlay?: boolean;
}): Record<string, string | boolean | undefined> {
  return {
    pos: options.position === undefined ? undefined : SIDES[options.position],
    align: options.alignment === undefined ? undefined : ALIGNMENTS[options.alignment],
    overlay: options.overlay,
  };
}

function parsePosition(el: XmlElement): {
  position?: ChartExSide;
  alignment?: ChartExAlignment;
  overlay?: boolean;
} {
  const position = attr(el, "pos") ?? attr(el, "cx:pos");
  const alignment = attr(el, "align") ?? attr(el, "cx:align");
  const overlay = boolean(el, "overlay", "positioned");
  return {
    ...(position === undefined ? {} : { position: SIDES_FROM[String(position)] }),
    ...(alignment === undefined ? {} : { alignment: ALIGNMENTS_FROM[String(alignment)] }),
    ...(overlay === undefined ? {} : { overlay }),
  };
}

function automaticXml(
  name: string,
  value: ChartExAutomaticNumber | undefined,
): Record<string, ChartExAutomaticNumber | undefined> {
  return value === undefined ? {} : { [name]: value };
}

function parseAutomatic(
  el: XmlElement,
  name: string,
  path: string,
): ChartExAutomaticNumber | undefined {
  const value = attr(el, name) ?? attr(el, `cx:${name}`);
  if (value === undefined) return undefined;
  return String(value) === "auto" ? "auto" : number(value, path, name);
}

function parseLevels(
  el: XmlElement,
  numeric: boolean,
): Array<ChartExNumericLevelLike | ChartExStringLevelLike> {
  const levels = children(el, "lvl").map((level, index) => {
    const path = `dimension/lvl[${index}]`;
    const points = children(level, "pt").map((point, pointIndex) => {
      const pointPath = `${path}/pt[${pointIndex}]`;
      const valueIndex = requiredNumber(point, "idx", pointPath);
      if (numeric) {
        const formatCode = attr(point, "formatCode");
        return {
          index: valueIndex,
          value: number(textOf(point), pointPath, "text"),
          ...(formatCode === undefined ? {} : { formatCode: String(formatCode) }),
        };
      }
      return { index: valueIndex, value: textOf(point) };
    });
    const name = attr(level, "name");
    const formatCode = attr(level, "formatCode");
    return {
      pointCount: requiredNumber(level, "ptCount", path),
      ...(name === undefined ? {} : { name: String(name) }),
      ...(formatCode === undefined ? {} : { formatCode: String(formatCode) }),
      ...(points.length ? { points } : {}),
    };
  });
  return levels as Array<ChartExNumericLevelLike | ChartExStringLevelLike>;
}

function levelsXml(
  levels: Array<ChartExNumericLevelLike | ChartExStringLevelLike>,
  numeric: boolean,
): string[] {
  return levels.map((level) =>
    element(
      "lvl",
      {
        ptCount: level.pointCount,
        name: "name" in level ? level.name : undefined,
        formatCode: "formatCode" in level ? level.formatCode : undefined,
      },
      (level.points ?? []).map((point) =>
        numeric
          ? element(
              "pt",
              { idx: point.index, formatCode: (point as ChartExNumericValue).formatCode },
              [String(point.value)],
            )
          : element("pt", { idx: point.index }, [(point as ChartExStringValue).value]),
      ),
    ),
  );
}

function parseDimension(el: XmlElement): ChartExDimension {
  if (!isName(el, "numDim") && !isName(el, "strDim"))
    throw new Error(`ChartEx unexpected dimension ${el.name}`);
  const path = `data/${el.name}`;
  const numeric = isName(el, "numDim");
  const type = required(el, "type", path) as ChartExDimension["type"];
  if (numeric !== NUMERIC_DIMENSIONS.has(type))
    throw new Error(`ChartEx ${path} dimension type does not match element`);
  const formula = parseFormula(findChild(el, "f"));
  const normalizedFormula = parseFormula(findChild(el, "nf"));
  const levels = parseLevels(el, numeric) as ChartExDimension["levels"];
  return {
    type,
    ...(formula ? { formula } : {}),
    ...(normalizedFormula ? { normalizedFormula } : {}),
    levels,
  } as ChartExDimension;
}

function dimensionXml(dimension: ChartExDimension): string {
  const numeric = NUMERIC_DIMENSIONS.has(dimension.type);
  return element(numeric ? "numDim" : "strDim", { type: dimension.type }, [
    formulaXml("f", dimension.formula),
    formulaXml("nf", dimension.normalizedFormula),
    ...levelsXml(dimension.levels, numeric),
  ]);
}

function parseData(el: XmlElement, index: number): ChartExData {
  const path = `data[${index}]`;
  return {
    id: requiredNumber(el, "id", path),
    dimensions: (el.elements ?? [])
      .filter((child) => isName(child, "numDim") || isName(child, "strDim"))
      .map(parseDimension),
    ...optionalExtensions(el),
  };
}

function chartDataXml(data: ChartExChartData, tagName = "chartData"): string {
  const external = data.externalData
    ? element("externalData", {
        "r:id": data.externalData.relationshipId,
        ...(data.externalData.autoUpdate === undefined
          ? {}
          : {
              [data.externalData.autoUpdatePrefix
                ? `${data.externalData.autoUpdatePrefix}:autoUpdate`
                : "autoUpdate"]: data.externalData.autoUpdate,
            }),
      })
    : "";
  return element(tagName, undefined, [
    external,
    ...data.data.map((item) =>
      element("data", { id: item.id }, [
        ...item.dimensions.map(dimensionXml),
        extensionsXml(item.extensions),
      ]),
    ),
    extensionsXml(data.extensions),
  ]);
}

function parseChartData(el: XmlElement | undefined): ChartExChartData {
  if (!el) return { data: [] };
  const external = findChild(el, "externalData");
  const autoUpdateAttribute = external
    ? Object.keys(external.attributes ?? {}).find(
        (name) => name === "autoUpdate" || name.endsWith(":autoUpdate"),
      )
    : undefined;
  return {
    ...(external
      ? {
          externalData: {
            relationshipId: required(external, "r:id", "chartData/externalData"),
            ...(autoUpdateAttribute === undefined
              ? {}
              : {
                  autoUpdate: boolean(external, autoUpdateAttribute, "chartData/externalData"),
                  ...(autoUpdateAttribute === "autoUpdate"
                    ? {}
                    : { autoUpdatePrefix: autoUpdateAttribute.slice(0, -":autoUpdate".length) }),
                }),
          },
        }
      : {}),
    data: children(el, "data").map(parseData),
    ...optionalExtensions(el),
  };
}

function axisXml(axis: ChartExAxis): string {
  const scaling =
    axis.scaling.kind === "category"
      ? element("catScaling", automaticXml("gapWidth", axis.scaling.gapWidth))
      : element("valScaling", {
          ...automaticXml("max", axis.scaling.maximum),
          ...automaticXml("min", axis.scaling.minimum),
          ...automaticXml("majorUnit", axis.scaling.majorUnit),
          ...automaticXml("minorUnit", axis.scaling.minorUnit),
        });
  const title = axis.title ? element("title", undefined, [containerXml(axis.title)]) : "";
  const units = axis.units
    ? element("units", { unit: axis.units.unit }, [containerXml(axis.units)])
    : "";
  const gridlines = (name: string, value?: ChartExGridlinesLike) =>
    value ? element(name, undefined, [containerXml(value)]) : "";
  const ticks = (name: string, value?: ChartExTickMarksLike) =>
    value
      ? element(name, { type: value.type === undefined ? undefined : TICK_MARKS[value.type] }, [
          extensionsXml(value.extensions),
        ])
      : "";
  return element("axis", { id: axis.id, hidden: axis.hidden }, [
    scaling,
    title,
    units,
    gridlines("majorGridlines", axis.majorGridlines),
    gridlines("minorGridlines", axis.minorGridlines),
    ticks("majorTickMarks", axis.majorTickMarks),
    ticks("minorTickMarks", axis.minorTickMarks),
    axis.tickLabels
      ? element("tickLabels", undefined, [extensionsXml(axis.tickLabels.extensions)])
      : "",
    numberFormatXml(axis.numberFormat),
    shapeXml(axis.shapeProperties),
    textPropertiesXml(axis.textProperties),
    extensionsXml(axis.extensions),
  ]);
}

function parseAxis(el: XmlElement, index: number): ChartExAxis {
  const path = `axis[${index}]`;
  const category = findChild(el, "catScaling");
  const value = findChild(el, "valScaling");
  let scaling: ChartExAxis["scaling"];
  if (category) {
    const gapWidth = parseAutomatic(category, "gapWidth", `${path}/catScaling`);
    scaling = { kind: "category", ...(gapWidth === undefined ? {} : { gapWidth }) };
  } else if (value) {
    const maximum = parseAutomatic(value, "max", `${path}/valScaling`);
    const minimum = parseAutomatic(value, "min", `${path}/valScaling`);
    const majorUnit = parseAutomatic(value, "majorUnit", `${path}/valScaling`);
    const minorUnit = parseAutomatic(value, "minorUnit", `${path}/valScaling`);
    scaling = {
      kind: "value",
      ...(maximum === undefined ? {} : { maximum }),
      ...(minimum === undefined ? {} : { minimum }),
      ...(majorUnit === undefined ? {} : { majorUnit }),
      ...(minorUnit === undefined ? {} : { minorUnit }),
    };
  } else {
    throw new Error(`ChartEx ${path} is missing scaling`);
  }
  const title = findChild(el, "title");
  const units = findChild(el, "units");
  const majorGridlines = findChild(el, "majorGridlines");
  const minorGridlines = findChild(el, "minorGridlines");
  const majorTickMarks = findChild(el, "majorTickMarks");
  const minorTickMarks = findChild(el, "minorTickMarks");
  const tickLabels = findChild(el, "tickLabels");
  const numberFormat = parseNumberFormat(findChild(el, "numFmt"), path);
  const shapeProperties = parseShape(findChild(el, "spPr"));
  const textProperties = parseTextProperties(findChild(el, "txPr"));
  const hidden = boolean(el, "hidden", path);
  return {
    id: requiredNumber(el, "id", path),
    scaling,
    ...(title ? { title: parseContainer(title) } : {}),
    ...(units
      ? {
          units: {
            unit: attr(units, "unit") as ChartExAxisUnit | undefined,
            ...parseContainer(units),
          },
        }
      : {}),
    ...(majorGridlines ? { majorGridlines: parseContainer(majorGridlines) } : {}),
    ...(minorGridlines ? { minorGridlines: parseContainer(minorGridlines) } : {}),
    ...(majorTickMarks
      ? {
          majorTickMarks: {
            type:
              attr(majorTickMarks, "type") === undefined
                ? undefined
                : TICK_MARKS_FROM[String(attr(majorTickMarks, "type"))],
            ...optionalExtensions(majorTickMarks),
          },
        }
      : {}),
    ...(minorTickMarks
      ? {
          minorTickMarks: {
            type:
              attr(minorTickMarks, "type") === undefined
                ? undefined
                : TICK_MARKS_FROM[String(attr(minorTickMarks, "type"))],
            ...optionalExtensions(minorTickMarks),
          },
        }
      : {}),
    ...(tickLabels ? { tickLabels: optionalExtensions(tickLabels) } : {}),
    ...(numberFormat ? { numberFormat } : {}),
    ...(shapeProperties ? { shapeProperties } : {}),
    ...(textProperties ? { textProperties } : {}),
    ...optionalExtensions(el),
    ...(hidden === undefined ? {} : { hidden }),
  };
}

function valueColorsXml(colors: ChartExValueColors | undefined): string {
  if (!colors) return "";
  const color = (name: string, value?: ChartExFill) =>
    value ? element(name, undefined, [fillDesc.stringify(value, {} as WriteContext) ?? ""]) : "";
  const positions = colors.positions;
  const position: string = positions
    ? element("valueColorPositions", { count: positions.count }, [
        positions.minimum ? colorChoiceXml(positions.minimum) : "",
        positions.middle ? middleChoiceXml(positions.middle) : "",
        positions.maximum ? colorChoiceXml(positions.maximum) : "",
      ])
    : "";
  return element("valueColors", undefined, [
    color("minColor", colors.minimum),
    color("midColor", colors.middle),
    color("maxColor", colors.maximum),
    position,
  ]);
}

function colorChoiceXml(
  value: ChartExValueColors["positions"] extends undefined
    ? never
    : NonNullable<NonNullable<ChartExValueColors["positions"]>["minimum"]>,
): string {
  if (value.extremeValue) return element("extremeValue");
  if (value.number !== undefined) return element("number", { val: value.number });
  if (value.percent !== undefined) return element("percent", { val: value.percent });
  throw new Error("ChartEx value color position is empty");
}

function middleChoiceXml(
  value: NonNullable<NonNullable<ChartExValueColors["positions"]>["middle"]>,
): string {
  if (value.number !== undefined) return element("number", { val: value.number });
  if (value.percent !== undefined) return element("percent", { val: value.percent });
  throw new Error("ChartEx middle value color position is empty");
}

function parseValueColors(el: XmlElement | undefined): ChartExValueColors | undefined {
  if (!el) return undefined;
  const color = (name: string): ChartExFill | undefined => {
    const child = findChild(el, name);
    return child ? fillDesc.parse(child, {} as ReadContext) : undefined;
  };
  const minimum = color("minColor");
  const middle = color("midColor");
  const maximum = color("maxColor");
  const positions = findChild(el, "valueColorPositions");
  return {
    ...(minimum || middle || maximum ? { minimum, middle, maximum } : {}),
    ...(positions
      ? {
          positions: {
            ...(attr(positions, "count") === undefined
              ? {}
              : { count: requiredNumber(positions, "count", "valueColorPositions") }),
            ...parseChoicePosition(findChild(positions, "min"), false),
            ...parseChoicePosition(findChild(positions, "mid"), true),
            ...parseChoicePosition(findChild(positions, "max"), false),
          },
        }
      : {}),
  };
}

function parseChoicePosition(
  el: XmlElement | undefined,
  middle: boolean,
): { minimum?: object; middle?: object; maximum?: object } {
  if (!el) return {};
  const extremeValue = !!findChild(el, "extremeValue");
  const numberElement = findChild(el, "number");
  const percentElement = findChild(el, "percent");
  const value = extremeValue
    ? { extremeValue: true as const }
    : numberElement
      ? { number: requiredNumber(numberElement, "val", "color position") }
      : percentElement
        ? { percent: requiredNumber(percentElement, "val", "color position") }
        : undefined;
  if (!value) throw new Error("ChartEx value color position is empty");
  return middle
    ? { middle: value as object }
    : extremeValue
      ? { minimum: value, maximum: value }
      : { minimum: value };
}

function visibilityXml(name: string, value: ChartExLabelVisibility | undefined): string {
  if (!value) return "";
  return element(name, {
    seriesName: value.seriesName,
    categoryName: value.categoryName,
    value: value.value,
  });
}

function parseLabelVisibility(el: XmlElement | undefined): ChartExLabelVisibility | undefined {
  if (!el) return undefined;
  const seriesName = boolean(el, "seriesName", "visibility");
  const categoryName = boolean(el, "categoryName", "visibility");
  const value = boolean(el, "value", "visibility");
  return {
    ...(seriesName === undefined ? {} : { seriesName }),
    ...(categoryName === undefined ? {} : { categoryName }),
    ...(value === undefined ? {} : { value }),
  };
}

function dataLabelsXml(labels: ChartExDataLabels | undefined): string {
  if (!labels) return "";
  const label = (item: ChartExDataLabel) =>
    element(
      "dataLabel",
      {
        idx: item.index,
        pos: item.position === undefined ? undefined : LABEL_POSITIONS[item.position],
      },
      [
        numberFormatXml(item.numberFormat),
        shapeXml(item.shapeProperties),
        textPropertiesXml(item.textProperties),
        visibilityXml("visibility", item.visibility),
        item.separator === undefined ? "" : element("separator", undefined, [item.separator]),
        extensionsXml(item.extensions),
      ],
    );
  return element(
    "dataLabels",
    { pos: labels.position === undefined ? undefined : LABEL_POSITIONS[labels.position] },
    [
      numberFormatXml(labels.numberFormat),
      shapeXml(labels.shapeProperties),
      textPropertiesXml(labels.textProperties),
      visibilityXml("visibility", labels.visibility),
      labels.separator === undefined ? "" : element("separator", undefined, [labels.separator]),
      ...(labels.labels ?? []).map(label),
      ...(labels.hidden ?? []).map((index) => element("dataLabelHidden", { idx: index })),
      extensionsXml(labels.extensions),
    ],
  );
}

function parseDataLabels(el: XmlElement | undefined): ChartExDataLabels | undefined {
  if (!el) return undefined;
  const position = attr(el, "pos");
  const numberFormat = parseNumberFormat(findChild(el, "numFmt"), "dataLabels");
  const shapeProperties = parseShape(findChild(el, "spPr"));
  const textProperties = parseTextProperties(findChild(el, "txPr"));
  const visibility = parseLabelVisibility(findChild(el, "visibility"));
  const separator = findChild(el, "separator");
  const extensions = parseExtensions(el);
  return {
    ...(position === undefined ? {} : { position: LABEL_POSITIONS_FROM[String(position)] }),
    ...(numberFormat ? { numberFormat } : {}),
    ...(shapeProperties ? { shapeProperties } : {}),
    ...(textProperties ? { textProperties } : {}),
    ...(visibility ? { visibility } : {}),
    ...(separator ? { separator: textOf(separator) } : {}),
    labels: children(el, "dataLabel").map((item) => {
      const itemPosition = attr(item, "pos");
      const itemFormat = parseNumberFormat(findChild(item, "numFmt"), "dataLabel");
      const itemShape = parseShape(findChild(item, "spPr"));
      const itemText = parseTextProperties(findChild(item, "txPr"));
      const itemVisibility = parseLabelVisibility(findChild(item, "visibility"));
      const itemSeparator = findChild(item, "separator");
      return {
        index: requiredNumber(item, "idx", "dataLabel"),
        ...(itemPosition === undefined
          ? {}
          : { position: LABEL_POSITIONS_FROM[String(itemPosition)] }),
        ...(itemFormat ? { numberFormat: itemFormat } : {}),
        ...(itemShape ? { shapeProperties: itemShape } : {}),
        ...(itemText ? { textProperties: itemText } : {}),
        ...(itemVisibility ? { visibility: itemVisibility } : {}),
        ...(itemSeparator ? { separator: textOf(itemSeparator) } : {}),
        ...optionalExtensions(item),
      };
    }),
    hidden: children(el, "dataLabelHidden").map((item) =>
      requiredNumber(item, "idx", "dataLabelHidden"),
    ),
    ...(extensions ? { extensions } : {}),
  };
}

function seriesLayoutPropertiesXml(properties: ChartExSeriesLayoutProperties | undefined): string {
  if (!properties) return "";
  const visibility = properties.visibility;
  const visibilityXml = visibility
    ? element("visibility", {
        connectorLines: visibility.connectorLines,
        meanLine: visibility.meanLine,
        meanMarker: visibility.meanMarker,
        nonoutliers: visibility.nonOutliers,
        outliers: visibility.outliers,
      })
    : "";
  const binning = properties.binning
    ? element(
        "binning",
        {
          intervalClosed:
            properties.binning.intervalClosed === undefined
              ? undefined
              : properties.binning.intervalClosed === "left"
                ? "l"
                : "r",
          underflow: properties.binning.underflow,
          overflow: properties.binning.overflow,
        },
        [
          properties.binning.size === undefined
            ? ""
            : element("binSize", { val: properties.binning.size }),
          properties.binning.count === undefined
            ? ""
            : element("binCount", { val: properties.binning.count }),
        ],
      )
    : "";
  const geography = properties.geography ? geographyXml(properties.geography) : "";
  const statistics = properties.statistics
    ? element("statistics", {
        quartileMethod:
          properties.statistics.quartileMethod === undefined
            ? undefined
            : properties.statistics.quartileMethod,
      })
    : "";
  const subtotals = properties.subtotals
    ? element(
        "subtotals",
        undefined,
        properties.subtotals.map((index) => element("idx", { val: index })),
      )
    : "";
  return element("layoutPr", undefined, [
    properties.parentLabelLayout
      ? element("parentLabelLayout", { val: properties.parentLabelLayout })
      : "",
    properties.regionLabelLayout
      ? element("regionLabelLayout", { val: properties.regionLabelLayout })
      : "",
    visibilityXml,
    properties.aggregation ? element("aggregation") : "",
    binning,
    geography,
    statistics,
    subtotals,
    extensionsXml(properties.extensions),
  ]);
}

function parseSeriesLayoutProperties(
  el: XmlElement | undefined,
): ChartExSeriesLayoutProperties | undefined {
  if (!el) return undefined;
  const parent = findChild(el, "parentLabelLayout");
  const region = findChild(el, "regionLabelLayout");
  const visibility = findChild(el, "visibility");
  const binning = findChild(el, "binning");
  const geography = findChild(el, "geography");
  const statistics = findChild(el, "statistics");
  const subtotals = findChild(el, "subtotals");
  const extensions = parseExtensions(el);
  if (parent && !["none", "banner", "overlapping"].includes(String(attr(parent, "val")))) {
    throw new Error("ChartEx parentLabelLayout has invalid value");
  }
  if (region && !["none", "bestFitOnly", "showAll"].includes(String(attr(region, "val")))) {
    throw new Error("ChartEx regionLabelLayout has invalid value");
  }
  return {
    ...(parent
      ? {
          parentLabelLayout: String(
            attr(parent, "val"),
          ) as ChartExSeriesLayoutProperties["parentLabelLayout"],
        }
      : {}),
    ...(region
      ? {
          regionLabelLayout: String(
            attr(region, "val"),
          ) as ChartExSeriesLayoutProperties["regionLabelLayout"],
        }
      : {}),
    ...(visibility
      ? {
          visibility: {
            ...(boolean(visibility, "connectorLines", "visibility") === undefined
              ? {}
              : { connectorLines: boolean(visibility, "connectorLines", "visibility") }),
            ...(boolean(visibility, "meanLine", "visibility") === undefined
              ? {}
              : { meanLine: boolean(visibility, "meanLine", "visibility") }),
            ...(boolean(visibility, "meanMarker", "visibility") === undefined
              ? {}
              : { meanMarker: boolean(visibility, "meanMarker", "visibility") }),
            ...(boolean(visibility, "nonoutliers", "visibility") === undefined
              ? {}
              : { nonOutliers: boolean(visibility, "nonoutliers", "visibility") }),
            ...(boolean(visibility, "outliers", "visibility") === undefined
              ? {}
              : { outliers: boolean(visibility, "outliers", "visibility") }),
          },
        }
      : {}),
    ...(findChild(el, "aggregation") ? { aggregation: true } : {}),
    ...(binning ? { binning: parseBinning(binning) } : {}),
    ...(geography ? { geography: parseGeography(geography) } : {}),
    ...(statistics
      ? {
          statistics: {
            quartileMethod: attr(statistics, "quartileMethod") as ChartExQuartileMethod | undefined,
          },
        }
      : {}),
    ...(subtotals
      ? {
          subtotals: children(subtotals, "idx").map((index) =>
            requiredNumber(index, "val", "subtotals/idx"),
          ),
        }
      : {}),
    ...(extensions ? { extensions } : {}),
  };
}

function parseBinning(el: XmlElement): ChartExBinning {
  const intervalClosed = attr(el, "intervalClosed");
  const underflow = parseAutomatic(el, "underflow", "binning");
  const overflow = parseAutomatic(el, "overflow", "binning");
  const sizeElement = findChild(el, "binSize");
  const countElement = findChild(el, "binCount");
  return {
    ...(intervalClosed === undefined
      ? {}
      : { intervalClosed: intervalClosed === "r" ? "right" : "left" }),
    ...(underflow === undefined ? {} : { underflow }),
    ...(overflow === undefined ? {} : { overflow }),
    ...(sizeElement ? { size: requiredNumber(sizeElement, "val", "binning/binSize") } : {}),
    ...(countElement ? { count: requiredNumber(countElement, "val", "binning/binCount") } : {}),
  };
}

function parseGeoCache(el: XmlElement): ChartExGeoCache {
  const entries: ChartExGeoCache["entries"] = [];
  for (const child of el.elements ?? []) {
    const localName = child.name?.includes(":")
      ? child.name.slice(child.name.indexOf(":") + 1)
      : child.name;
    if (localName === "binary") entries.push({ kind: "binary", value: textOf(child) });
    else if (isName(child, "clear")) entries.push({ kind: "clear" });
    else throw new Error(`ChartEx geoCache has unexpected child ${child.name}`);
  }
  return { provider: required(el, "provider", "geoCache"), entries };
}

function parseGeography(el: XmlElement): ChartExGeography {
  const cache = findChild(el, "geoCache");
  return {
    cultureLanguage: required(el, "cultureLanguage", "geography"),
    cultureRegion: required(el, "cultureRegion", "geography"),
    attribution: required(el, "attribution", "geography"),
    ...(attr(el, "projectionType") === undefined
      ? {}
      : { projection: attr(el, "projectionType") as ChartExGeography["projection"] }),
    ...(attr(el, "viewedRegionType") === undefined
      ? {}
      : { viewedRegion: attr(el, "viewedRegionType") as ChartExGeography["viewedRegion"] }),
    ...(cache ? { cache: parseGeoCache(cache) } : {}),
  };
}

function geoCacheXml(cache: ChartExGeoCache): string {
  return element(
    "geoCache",
    { provider: cache.provider },
    cache.entries.map((entry) => {
      if (entry.kind === "binary") return element("binary", undefined, [entry.value]);
      return element("clear");
    }),
  );
}

function geographyXml(geography: ChartExGeography): string {
  return element(
    "geography",
    {
      projectionType: geography.projection,
      viewedRegionType: geography.viewedRegion,
      cultureLanguage: geography.cultureLanguage,
      cultureRegion: geography.cultureRegion,
      attribution: geography.attribution,
    },
    [geography.cache ? geoCacheXml(geography.cache) : ""],
  );
}

function seriesXml(series: ChartExSeries): string {
  const dataId = series.dataId === undefined ? "" : element("dataId", { val: series.dataId });
  return element(
    "series",
    {
      layoutId: series.layout,
      hidden: series.hidden,
      ownerIdx: series.ownerIndex,
      uniqueId: series.uniqueId,
      formatIdx: series.formatIndex,
    },
    [
      textXml(series.text),
      shapeXml(series.shapeProperties),
      valueColorsXml(series.colors),
      ...(series.dataPoints ?? []).map((point) =>
        element("dataPt", { idx: point.index }, [
          shapeXml(point.shapeProperties),
          extensionsXml(point.extensions),
        ]),
      ),
      dataLabelsXml(series.dataLabels),
      dataId,
      seriesLayoutPropertiesXml(series.layoutProperties),
      ...(series.axisIds ?? []).map((id) => element("axisId", { val: id })),
      extensionsXml(series.extensions),
    ],
  );
}

function parseSeries(el: XmlElement): ChartExSeries {
  const path = "series";
  const text = parseText(findChild(el, "tx"));
  const shapeProperties = parseShape(findChild(el, "spPr"));
  const colors = parseValueColors(findChild(el, "valueColors"));
  const dataLabels = parseDataLabels(findChild(el, "dataLabels"));
  const dataIdElement = findChild(el, "dataId");
  const layoutProperties = parseSeriesLayoutProperties(findChild(el, "layoutPr"));
  const extensions = parseExtensions(el);
  const hidden = boolean(el, "hidden", path);
  const ownerIndex = attr(el, "ownerIdx");
  const uniqueId = attr(el, "uniqueId");
  const formatIndex = attr(el, "formatIdx");
  return {
    layout: required(el, "layoutId", path) as ChartExSeries["layout"],
    ...(text ? { text } : {}),
    ...(shapeProperties ? { shapeProperties } : {}),
    ...(colors ? { colors } : {}),
    dataPoints: children(el, "dataPt").map((point) => {
      const pointShape = parseShape(findChild(point, "spPr"));
      return {
        index: requiredNumber(point, "idx", "dataPt"),
        ...(pointShape ? { shapeProperties: pointShape } : {}),
        ...optionalExtensions(point),
      };
    }),
    ...(dataLabels ? { dataLabels } : {}),
    ...(dataIdElement ? { dataId: requiredNumber(dataIdElement, "val", "series/dataId") } : {}),
    ...(layoutProperties ? { layoutProperties } : {}),
    axisIds: children(el, "axisId").map((axisId) =>
      number(attr(axisId, "val") ?? textOf(axisId), "series/axisId", "value"),
    ),
    ...(extensions ? { extensions } : {}),
    ...(hidden === undefined ? {} : { hidden }),
    ...(ownerIndex === undefined ? {} : { ownerIndex: number(ownerIndex, path, "ownerIdx") }),
    ...(uniqueId === undefined ? {} : { uniqueId: String(uniqueId) }),
    ...(formatIndex === undefined ? {} : { formatIndex: number(formatIndex, path, "formatIdx") }),
  };
}

function chartXml(chart: ChartExChart): string {
  return element("chart", undefined, [
    chart.title ? element("title", positionXml(chart.title), [containerXml(chart.title)]) : "",
    element("plotArea", undefined, [
      element("plotAreaRegion", undefined, [
        chart.plotArea.region.surface
          ? element("plotSurface", undefined, [containerXml(chart.plotArea.region.surface)])
          : "",
        ...chart.plotArea.region.series.map(seriesXml),
        extensionsXml(chart.plotArea.region.extensions),
      ]),
      ...chart.plotArea.axes.map(axisXml),
      shapeXml(chart.plotArea.shapeProperties),
      extensionsXml(chart.plotArea.extensions),
    ]),
    chart.legend ? element("legend", positionXml(chart.legend), [containerXml(chart.legend)]) : "",
    extensionsXml(chart.extensions),
  ]);
}

function parseChart(el: XmlElement): ChartExChart {
  const title = findChild(el, "title");
  const legend = findChild(el, "legend");
  const plotArea = findChild(el, "plotArea");
  if (!plotArea) throw new Error("ChartEx chart is missing plotArea");
  const region = findChild(plotArea, "plotAreaRegion");
  if (!region) throw new Error("ChartEx plotArea is missing plotAreaRegion");
  const surface = findChild(region, "plotSurface");
  const regionExtensions = parseExtensions(region);
  const plotShape = parseShape(findChild(plotArea, "spPr"));
  const plotExtensions = parseExtensions(plotArea);
  return {
    ...(title ? { title: { ...parsePosition(title), ...parseContainer(title) } } : {}),
    plotArea: {
      region: {
        ...(surface ? { surface: parseContainer(surface) } : {}),
        series: children(region, "series").map(parseSeries),
        ...(regionExtensions ? { extensions: regionExtensions } : {}),
      },
      axes: children(plotArea, "axis").map(parseAxis),
      ...(plotShape ? { shapeProperties: plotShape } : {}),
      ...(plotExtensions ? { extensions: plotExtensions } : {}),
    },
    ...(legend ? { legend: { ...parsePosition(legend), ...parseContainer(legend) } } : {}),
    ...optionalExtensions(el),
  };
}

function formatOverridesXml(overrides: ChartExFormatOverride[] | undefined): string[] {
  return (overrides ?? []).map((override) =>
    element("fmtOvr", { idx: override.index }, [
      shapeXml(override.shapeProperties),
      extensionsXml(override.extensions),
    ]),
  );
}

function parseFormatOverrides(el: XmlElement): ChartExFormatOverride[] {
  return children(el, "fmtOvr").map((override) => {
    const shapeProperties = parseShape(findChild(override, "spPr"));
    return {
      index: requiredNumber(override, "idx", "fmtOvrs/fmtOvr"),
      ...(shapeProperties ? { shapeProperties } : {}),
      ...optionalExtensions(override),
    };
  });
}

function printSettingsXml(settings: ChartExPrintSettings | undefined): string {
  if (!settings) return "";
  const headerFooter = settings.headerFooter
    ? element(
        "headerFooter",
        {
          alignWithMargins: settings.headerFooter.alignWithMargins,
          differentOddEven: settings.headerFooter.differentOddEven,
          differentFirst: settings.headerFooter.differentFirst,
        },
        [
          textChild("oddHeader", settings.headerFooter.oddHeader),
          textChild("oddFooter", settings.headerFooter.oddFooter),
          textChild("evenHeader", settings.headerFooter.evenHeader),
          textChild("evenFooter", settings.headerFooter.evenFooter),
          textChild("firstHeader", settings.headerFooter.firstHeader),
          textChild("firstFooter", settings.headerFooter.firstFooter),
        ],
      )
    : "";
  const pageMargins = settings.pageMargins
    ? element("pageMargins", { ...settings.pageMargins })
    : "";
  const pageSetup = settings.pageSetup ? element("pageSetup", { ...settings.pageSetup }) : "";
  return element("printSettings", undefined, [headerFooter, pageMargins, pageSetup]);
}

function textChild(name: string, value: string | undefined): string {
  return value === undefined ? "" : element(name, undefined, [value]);
}

function parsePrintSettings(el: XmlElement | undefined): ChartExPrintSettings | undefined {
  if (!el) return undefined;
  const headerFooterElement = findChild(el, "headerFooter");
  const pageMarginsElement = findChild(el, "pageMargins");
  const pageSetupElement = findChild(el, "pageSetup");
  return {
    ...(headerFooterElement
      ? {
          headerFooter: {
            oddHeader: textValue(headerFooterElement, "oddHeader"),
            oddFooter: textValue(headerFooterElement, "oddFooter"),
            evenHeader: textValue(headerFooterElement, "evenHeader"),
            evenFooter: textValue(headerFooterElement, "evenFooter"),
            firstHeader: textValue(headerFooterElement, "firstHeader"),
            firstFooter: textValue(headerFooterElement, "firstFooter"),
            ...(boolean(headerFooterElement, "alignWithMargins", "headerFooter") === undefined
              ? {}
              : {
                  alignWithMargins: boolean(
                    headerFooterElement,
                    "alignWithMargins",
                    "headerFooter",
                  ),
                }),
            ...(boolean(headerFooterElement, "differentOddEven", "headerFooter") === undefined
              ? {}
              : {
                  differentOddEven: boolean(
                    headerFooterElement,
                    "differentOddEven",
                    "headerFooter",
                  ),
                }),
            ...(boolean(headerFooterElement, "differentFirst", "headerFooter") === undefined
              ? {}
              : { differentFirst: boolean(headerFooterElement, "differentFirst", "headerFooter") }),
          },
        }
      : {}),
    ...(pageMarginsElement
      ? {
          pageMargins: {
            left: requiredNumber(pageMarginsElement, "l", "pageMargins"),
            right: requiredNumber(pageMarginsElement, "r", "pageMargins"),
            top: requiredNumber(pageMarginsElement, "t", "pageMargins"),
            bottom: requiredNumber(pageMarginsElement, "b", "pageMargins"),
            header: requiredNumber(pageMarginsElement, "header", "pageMargins"),
            footer: requiredNumber(pageMarginsElement, "footer", "pageMargins"),
          },
        }
      : {}),
    ...(pageSetupElement
      ? { pageSetup: { ...pageSetupElement.attributes } as ChartExPageSetup }
      : {}),
  };
}

function textValue(parent: XmlElement, name: string): string | undefined {
  const child = findChild(parent, name);
  return child ? textOf(child) : undefined;
}

export const chartExSpaceDesc: CustomDescriptor<ChartExSpaceOptions> = {
  kind: "custom",

  stringify(options, ctx) {
    void ctx;
    if (options.chart === undefined && options.dashboard === undefined) {
      throw new Error("ChartEx chartSpace requires chart or dashboard");
    }
    return element(
      "chartSpace",
      {
        version: options.version,
        featureList: options.featureList,
        fallbackImg: options.fallbackImage,
      },
      [
        options.chartData ? chartDataXml(options.chartData) : "",
        options.chartDataIntentionallyChanged
          ? chartDataXml(options.chartDataIntentionallyChanged, "chartDataIntentionallyChanged")
          : "",
        options.chart ? chartXml(options.chart) : dashboardXml(options.dashboard),
        shapeXml(options.shapeProperties),
        textPropertiesXml(options.textProperties),
        options.colorMapping ? stringifyColorMapping(options.colorMapping, "cx:clrMapOvr") : "",
        options.formatOverrides
          ? element("fmtOvrs", undefined, formatOverridesXml(options.formatOverrides))
          : "",
        printSettingsXml(options.printSettings),
        extensionsXml(options.extensions),
      ],
    );
  },

  parse(el, ctx) {
    void ctx;
    if (!isName(el, "chartSpace")) throw new Error("ChartEx root must be chartSpace");
    const chart = findChild(el, "chart");
    const dashboard = findChild(el, "dashboard");
    if (!chart && !dashboard) throw new Error("ChartEx chartSpace requires chart or dashboard");
    const knownChildren = new Set([
      "chartData",
      "chartDataIntentionallyChanged",
      "chart",
      "dashboard",
      "spPr",
      "txPr",
      "clrMapOvr",
      "fmtOvrs",
      "printSettings",
      "extLst",
    ]);
    for (const [index, child] of (el.elements ?? []).entries()) {
      const name = child.name?.includes(":")
        ? child.name.slice(child.name.indexOf(":") + 1)
        : child.name;
      if (!knownChildren.has(name ?? "")) {
        throw new ChartExParseError(
          `ChartEx chartSpace has unknown child ${child.name}`,
          "chartSpace",
          `/chartSpace[${index}]`,
          "unknown-chartex-child",
        );
      }
    }
    const shapeProperties = parseShape(findChild(el, "spPr"));
    const textProperties = parseTextProperties(findChild(el, "txPr"));
    const formatOverridesElement = findChild(el, "fmtOvrs");
    const printSettings = parsePrintSettings(findChild(el, "printSettings"));
    const extensions = parseExtensions(el);
    const version = attr(el, "version");
    const featureList = attr(el, "featureList");
    const fallbackImage = attr(el, "fallbackImg");
    return {
      ...(findChild(el, "chartData")
        ? { chartData: parseChartData(findChild(el, "chartData")) }
        : {}),
      ...(findChild(el, "chartDataIntentionallyChanged")
        ? {
            chartDataIntentionallyChanged: parseChartData(
              findChild(el, "chartDataIntentionallyChanged"),
            ),
          }
        : {}),
      ...(chart ? { chart: parseChart(chart) } : {}),
      ...(dashboard ? { dashboard: parseDashboard(dashboard) } : {}),
      ...(shapeProperties ? { shapeProperties } : {}),
      ...(textProperties ? { textProperties } : {}),
      ...(parseColorMapping(findChild(el, "cx:clrMapOvr"))
        ? { colorMapping: parseColorMapping(findChild(el, "cx:clrMapOvr")) }
        : {}),
      ...(formatOverridesElement
        ? { formatOverrides: parseFormatOverrides(formatOverridesElement) }
        : {}),
      ...(printSettings ? { printSettings } : {}),
      ...(extensions ? { extensions } : {}),
      ...(version === undefined ? {} : { version: String(version) }),
      ...(featureList === undefined ? {} : { featureList: String(featureList) }),
      ...(fallbackImage === undefined ? {} : { fallbackImage: String(fallbackImage) }),
    };
  },
};

function dashboardXml(dashboard: ChartExDashboard | undefined): string {
  if (!dashboard) throw new Error("ChartEx chartSpace requires chart or dashboard");
  return element("dashboard", { entrypoint: dashboard.entrypoint }, [
    element(
      "bindings",
      undefined,
      dashboard.bindings.map((binding) =>
        element("binding", { name: binding.name, id: binding.dataId, type: binding.type }),
      ),
    ),
    element(
      "artifacts",
      undefined,
      dashboard.artifacts.map((artifact) =>
        element("artifact", { source: artifact.source, name: artifact.name, type: artifact.type }, [
          artifact.value,
        ]),
      ),
    ),
    extensionsXml(dashboard.extensions),
  ]);
}

function parseDashboard(el: XmlElement): ChartExDashboard {
  const bindings = findChild(el, "bindings");
  const artifacts = findChild(el, "artifacts");
  if (!bindings || !artifacts) throw new Error("ChartEx dashboard requires bindings and artifacts");
  const extensions = parseExtensions(el);
  return {
    entrypoint: required(el, "entrypoint", "dashboard"),
    bindings: children(bindings, "binding").map((binding) => ({
      name: required(binding, "name", "binding"),
      dataId: requiredNumber(binding, "id", "binding"),
      type: required(binding, "type", "binding") as ChartExBinding["type"],
    })),
    artifacts: children(artifacts, "artifact").map((artifact) => {
      const source = attr(artifact, "source");
      return {
        name: required(artifact, "name", "artifact"),
        type: required(artifact, "type", "artifact") as "html",
        value: textOf(artifact),
        ...(source === undefined ? {} : { source: String(source) }),
      };
    }),
    ...(extensions ? { extensions } : {}),
  };
}
