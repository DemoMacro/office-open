/**
 * Canonical ChartEx model (Microsoft chartex extension parts).
 *
 * @module
 */

import type { FillOptions } from "../drawing/fill/fill-options";
import type { ShapePropertiesOptions } from "../drawing/shape-properties-desc";
import type { TextBodyOptions } from "../drawing/text/text-body";
import type { ColorMappingOptions } from "../theme/theme-options";

export type ChartExAutomatic = "auto";
export type ChartExAutomaticNumber = number | ChartExAutomatic;
export type ChartExSide = "left" | "top" | "right" | "bottom";
export type ChartExAlignment = "minimum" | "center" | "maximum";
export type ChartExFormulaDirection = "column" | "row";
export type ChartExStringDimensionType = "category" | "colorString" | "entityId";
export type ChartExNumericDimensionType =
  | "val"
  | "x"
  | "xVal"
  | "y"
  | "yVal"
  | "size"
  | "degree"
  | "colorVal";
export type ChartExAxisUnit =
  | "hundreds"
  | "thousands"
  | "tenThousands"
  | "hundredThousands"
  | "millions"
  | "tenMillions"
  | "hundredMillions"
  | "billions"
  | "trillions"
  | "percentage";
export type ChartExTickMarkType = "inside" | "outside" | "cross" | "none";
export type ChartExSeriesLayout =
  | "boxWhisker"
  | "clusteredColumn"
  | "funnel"
  | "paretoLine"
  | "regionMap"
  | "sunburst"
  | "treemap"
  | "waterfall";
export type ChartExLabelLayout = "none" | "banner" | "overlapping";
export type ChartExRegionLabelLayout = "none" | "bestFitOnly" | "showAll";
export type ChartExIntervalClosedSide = "left" | "right";
export type ChartExEntityType =
  | "Address"
  | "AdminDistrict"
  | "AdminDistrict2"
  | "AdminDistrict3"
  | "Continent"
  | "CountryRegion"
  | "Locality"
  | "Ocean"
  | "Planet"
  | "PostalCode"
  | "Region"
  | "Unsupported";
export type ChartExProjection = "mercator" | "miller" | "robinson" | "albers";
export type ChartExMappingLevel =
  | "dataOnly"
  | "postalCode"
  | "county"
  | "state"
  | "countryRegion"
  | "countryRegionList"
  | "world";
export type ChartExQuartileMethod = "inclusive" | "exclusive";
export type ChartExLabelPosition =
  | "bestFit"
  | "bottom"
  | "center"
  | "insideBase"
  | "insideEnd"
  | "left"
  | "outsideEnd"
  | "right"
  | "top";

export interface ChartExExtension {
  uri?: string;
  content?: string;
}

export class ChartExParseError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly reason: string,
  ) {
    super(message);
    this.name = "ChartExParseError";
  }
}

export interface ChartExNumberFormat {
  formatCode: string;
  sourceLinked?: boolean;
}

export interface ChartExFormula {
  text: string;
  direction?: ChartExFormulaDirection;
}

export interface ChartExExternalData {
  relationshipId: string;
  autoUpdate?: boolean;
  /** Attribute-name prefix for @autoUpdate; round-trip only. */
  autoUpdatePrefix?: string;
}

export interface ChartExStringValue {
  index: number;
  value: string;
}

export interface ChartExStringLevel {
  pointCount: number;
  name?: string;
  points?: ChartExStringValue[];
}

export interface ChartExStringDimension {
  type: ChartExStringDimensionType;
  formula?: ChartExFormula;
  normalizedFormula?: ChartExFormula;
  levels: ChartExStringLevel[];
}

export interface ChartExNumericValue {
  index: number;
  value: number;
  formatCode?: string;
}

export interface ChartExNumericLevel {
  pointCount: number;
  formatCode?: string;
  points?: ChartExNumericValue[];
}

export interface ChartExNumericDimension {
  type: ChartExNumericDimensionType;
  formula?: ChartExFormula;
  normalizedFormula?: ChartExFormula;
  levels: ChartExNumericLevel[];
}

export type ChartExDimension = ChartExNumericDimension | ChartExStringDimension;

export interface ChartExData {
  id: number;
  dimensions: ChartExDimension[];
  extensions?: ChartExExtension[];
}

export interface ChartExChartData {
  externalData?: ChartExExternalData;
  data: ChartExData[];
  extensions?: ChartExExtension[];
}

export type ChartExText = { formula?: ChartExFormula; value?: string } | { rich: TextBodyOptions };

export interface ChartExOffset {
  top: number;
  left: number;
}

interface ChartExPositionedProperties {
  position?: ChartExSide;
  alignment?: ChartExAlignment;
  overlay?: boolean;
}

export interface ChartExTitle extends ChartExPositionedProperties {
  text?: ChartExText;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  offset?: ChartExOffset;
  extensions?: ChartExExtension[];
}

export interface ChartExCategoryScaling {
  kind: "category";
  gapWidth?: ChartExAutomaticNumber;
}

export interface ChartExValueScaling {
  kind: "value";
  maximum?: ChartExAutomaticNumber;
  minimum?: ChartExAutomaticNumber;
  majorUnit?: ChartExAutomaticNumber;
  minorUnit?: ChartExAutomaticNumber;
}

export type ChartExScaling = ChartExCategoryScaling | ChartExValueScaling;

export interface ChartExUnitsLabel {
  text?: ChartExText;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  extensions?: ChartExExtension[];
}

export interface ChartExUnits {
  unit?: ChartExAxisUnit;
  label?: ChartExUnitsLabel;
  extensions?: ChartExExtension[];
}

export interface ChartExGridlines {
  shapeProperties?: ShapePropertiesOptions;
  extensions?: ChartExExtension[];
}

export interface ChartExTickMarks {
  type?: ChartExTickMarkType;
  extensions?: ChartExExtension[];
}

export interface ChartExTickLabels {
  extensions?: ChartExExtension[];
}

export interface ChartExAxis {
  id: number;
  scaling: ChartExScaling;
  title?: Omit<ChartExTitle, keyof ChartExPositionedProperties>;
  units?: ChartExUnits;
  majorGridlines?: ChartExGridlines;
  minorGridlines?: ChartExGridlines;
  majorTickMarks?: ChartExTickMarks;
  minorTickMarks?: ChartExTickMarks;
  tickLabels?: ChartExTickLabels;
  numberFormat?: ChartExNumberFormat;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  extensions?: ChartExExtension[];
  hidden?: boolean;
}

export interface ChartExValueColors {
  minimum?: FillOptions;
  middle?: FillOptions;
  maximum?: FillOptions;
  positions?: {
    count?: number;
    minimum?: { extremeValue?: true; number?: number; percent?: number };
    middle?: { number?: number; percent?: number };
    maximum?: { extremeValue?: true; number?: number; percent?: number };
  };
}

export interface ChartExLabelVisibility {
  seriesName?: boolean;
  categoryName?: boolean;
  value?: boolean;
}

export interface ChartExSeriesVisibility {
  connectorLines?: boolean;
  meanLine?: boolean;
  meanMarker?: boolean;
  nonOutliers?: boolean;
  outliers?: boolean;
}

export interface ChartExGeoCache {
  provider: string;
  entries: Array<
    | { kind: "binary"; value: string }
    | {
        kind: "clear";
        locationQueries?: Array<{
          query?: {
            countryRegion?: string;
            adminDistrict1?: string;
            adminDistrict2?: string;
            postalCode?: string;
            entityType: ChartExEntityType;
          };
          locations?: Array<{
            address?: {
              address1?: string;
              countryRegion?: string;
              adminDistrict1?: string;
              postalCode?: string;
              locality?: string;
              isoCountryCode?: string;
            };
            latitude?: number;
            longitude?: number;
            entityName: string;
            entityType: ChartExEntityType;
          }>;
        }>;
        entityQueries?: Array<{
          query?: { entityType: ChartExEntityType; entityId: string };
          data?: {
            entityName: string;
            entityId: string;
            east: number;
            west: number;
            north: number;
            south: number;
            polygons?: Array<{ polygonId: string; pointCount: number; pcaRings: string }>;
            copyrights?: string[];
          };
        }>;
        pointToEntityQueries?: Array<{
          query: { entityType: ChartExEntityType; entityId: string };
        }>;
        childQueries?: Array<{ entityId: string; entityTypes?: ChartExEntityType[] }>;
        parentQueries?: Array<{ entityId: string }>;
      }
  >;
}

export interface ChartExGeography {
  cultureLanguage: string;
  cultureRegion: string;
  attribution: string;
  projection?: ChartExProjection;
  viewedRegion?: ChartExMappingLevel;
  cache?: ChartExGeoCache;
}

export interface ChartExBinning {
  intervalClosed?: ChartExIntervalClosedSide;
  underflow?: ChartExAutomaticNumber;
  overflow?: ChartExAutomaticNumber;
  size?: number;
  count?: number;
}

export interface ChartExSeriesLayoutProperties {
  parentLabelLayout?: ChartExLabelLayout;
  regionLabelLayout?: ChartExRegionLabelLayout;
  visibility?: ChartExSeriesVisibility;
  aggregation?: true;
  binning?: ChartExBinning;
  geography?: ChartExGeography;
  statistics?: { quartileMethod?: ChartExQuartileMethod };
  subtotals?: number[];
  extensions?: ChartExExtension[];
}

export interface ChartExDataPoint {
  index: number;
  shapeProperties?: ShapePropertiesOptions;
  extensions?: ChartExExtension[];
}

export interface ChartExDataLabel {
  index: number;
  position?: ChartExLabelPosition;
  numberFormat?: ChartExNumberFormat;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  visibility?: ChartExLabelVisibility;
  separator?: string;
  extensions?: ChartExExtension[];
}

export interface ChartExDataLabels {
  position?: ChartExLabelPosition;
  numberFormat?: ChartExNumberFormat;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  visibility?: ChartExLabelVisibility;
  separator?: string;
  labels?: ChartExDataLabel[];
  hidden?: number[];
  extensions?: ChartExExtension[];
}

export interface ChartExSeries {
  layout: ChartExSeriesLayout;
  text?: ChartExText;
  shapeProperties?: ShapePropertiesOptions;
  colors?: ChartExValueColors;
  dataPoints?: ChartExDataPoint[];
  dataLabels?: ChartExDataLabels;
  dataId?: number;
  layoutProperties?: ChartExSeriesLayoutProperties;
  axisIds?: number[];
  extensions?: ChartExExtension[];
  hidden?: boolean;
  ownerIndex?: number;
  uniqueId?: string;
  formatIndex?: number;
}

export interface ChartExLegend extends ChartExPositionedProperties {
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  offset?: ChartExOffset;
  extensions?: ChartExExtension[];
}

export interface ChartExPlotSurface {
  shapeProperties?: ShapePropertiesOptions;
  extensions?: ChartExExtension[];
}

export interface ChartExPlotArea {
  region: {
    surface?: ChartExPlotSurface;
    series: ChartExSeries[];
    extensions?: ChartExExtension[];
  };
  axes: ChartExAxis[];
  shapeProperties?: ShapePropertiesOptions;
  extensions?: ChartExExtension[];
}

export interface ChartExChart {
  title?: ChartExTitle;
  plotArea: ChartExPlotArea;
  legend?: ChartExLegend;
  extensions?: ChartExExtension[];
}

export interface ChartExBinding {
  name: string;
  dataId: number;
  type: ChartExNumericDimensionType | ChartExStringDimensionType;
}

export interface ChartExArtifact {
  name: string;
  type: "html";
  value: string;
  source?: string;
}

export interface ChartExDashboard {
  entrypoint: string;
  bindings: ChartExBinding[];
  artifacts: ChartExArtifact[];
  extensions?: ChartExExtension[];
}

export interface ChartExFormatOverride {
  index: number;
  shapeProperties?: ShapePropertiesOptions;
  extensions?: ChartExExtension[];
}

export interface ChartExHeaderFooter {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  alignWithMargins?: boolean;
  differentOddEven?: boolean;
  differentFirst?: boolean;
}

export interface ChartExPageMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

export type ChartExPageOrientation = "default" | "portrait" | "landscape";

export interface ChartExPageSetup {
  paperSize?: number;
  firstPageNumber?: number;
  orientation?: ChartExPageOrientation;
  blackAndWhite?: boolean;
  draft?: boolean;
  useFirstPageNumber?: boolean;
  horizontalDpi?: number;
  verticalDpi?: number;
  copies?: number;
}

export interface ChartExPrintSettings {
  headerFooter?: ChartExHeaderFooter;
  pageMargins?: ChartExPageMargins;
  pageSetup?: ChartExPageSetup;
}

export interface ChartExSpaceOptions {
  chartData?: ChartExChartData;
  chart?: ChartExChart;
  dashboard?: ChartExDashboard;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
  colorMapping?: Partial<ColorMappingOptions>;
  formatOverrides?: ChartExFormatOverride[];
  printSettings?: ChartExPrintSettings;
  extensions?: ChartExExtension[];
  version?: string;
  featureList?: string;
  fallbackImage?: string;
}
