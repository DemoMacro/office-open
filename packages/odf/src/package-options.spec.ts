import { parse } from "@office-open/xml";
import { describe, expect, it } from "vitest";

import {
  chartPropertiesXml,
  listLevelPropertiesXml,
  parseChartStyleProperties,
  parseListLevelProperties,
  parseStyleProperty,
  stylePropertyXml,
  type ChartStylePropertiesOptions,
  type ListLevelStylePropertiesOptions,
  type StylePropertyOptions,
} from "./package-options";

function element(xml: string) {
  const result = parse(xml, { ignoreDeclaration: true }).elements?.at(0);
  if (!result) throw new Error("Test XML has no element");
  return result;
}

function property(xml: string): StylePropertyOptions {
  return propertyIn(xml, "paragraph");
}

function propertyIn(
  xml: string,
  family: "graphic" | "page-layout" | "drawing-page" | "paragraph" | "section" | "table-cell",
) {
  const parsed = parseStyleProperty(element(xml), { family, scale: "emu" });
  if (!parsed.ok)
    throw new Error(
      `${family}: ${parsed.elementName}: ${parsed.attributeName ?? ""}: ${parsed.reason}`,
    );
  return parsed.value;
}

describe("ODF style property codecs", () => {
  it("round-trips graphic and text glow extensions", () => {
    const graphic: StylePropertyOptions = {
      name: "style:graphic-properties",
      attributes: {
        "loext:glow-radius": 3810000,
        "loext:glow-color": "#e97132",
        "loext:glow-transparency": 5,
      },
    };
    expect(
      parseStyleProperty(element(stylePropertyXml(graphic, "emu")), {
        family: "graphic",
        scale: "emu",
      }),
    ).toEqual({ ok: true, value: graphic });
    const text: StylePropertyOptions = {
      name: "style:text-properties",
      attributes: {
        "loext:glow-text-radius": 2540000,
        "loext:glow-text-color": "#4ea72e",
        "loext:glow-text-transparency": 5,
      },
    };
    expect(
      parseStyleProperty(element(stylePropertyXml(text, "emu")), {
        family: "text",
        scale: "emu",
      }),
    ).toEqual({ ok: true, value: text });
    const cell: StylePropertyOptions = {
      name: "style:table-cell-properties",
      attributes: { "fo:text-align": "center", "style:width": 914400 },
    };
    expect(
      parseStyleProperty(element(stylePropertyXml(cell, "emu")), {
        family: "table-cell",
        scale: "emu",
      }),
    ).toEqual({ ok: true, value: cell });
  });

  it("round-trips chart properties and symbol images", () => {
    const options: ChartStylePropertiesOptions = {
      name: "style:chart-properties",
      attributes: {
        "chart:vertical": true,
        "chart:gap-width": 75,
        "chart:regression-max-degree": 3,
        "chart:error-category": "standard-deviation",
        "chart:interval-major": 2.5,
        "chart:symbol-type": "image",
      },
      symbolImage: { attributes: { "xlink:href": "media/symbol.png" } },
    };
    expect(parseChartStyleProperties(element(chartPropertiesXml(options, "emu")), "emu")).toEqual({
      ok: true,
      value: options,
    });
  });

  it("round-trips every chart property attribute type", () => {
    const options: ChartStylePropertiesOptions = {
      name: "style:chart-properties",
      attributes: {
        "chart:angle-offset": 45,
        "chart:auto-position": true,
        "chart:auto-size": false,
        "chart:axis-label-position": "outside",
        "chart:axis-position": 0,
        "chart:connect-bars": true,
        "chart:data-label-number": "value",
        "chart:data-label-series": true,
        "chart:data-label-symbol": false,
        "chart:data-label-text": true,
        "chart:deep": true,
        "chart:display-label": true,
        "chart:error-category": "constant",
        "chart:error-lower-indicator": true,
        "chart:error-lower-limit": 1.5,
        "chart:error-lower-range": "Sheet1.A1:A5",
        "chart:error-margin": 0.5,
        "chart:error-percentage": 10,
        "chart:error-upper-indicator": true,
        "chart:error-upper-limit": 2.5,
        "chart:error-upper-range": "Sheet1.B1:B5",
        "chart:gap-width": 75,
        "chart:group-bars-per-axis": true,
        "chart:hole-size": 20,
        "chart:include-hidden-cells": true,
        "chart:interpolation": "none",
        "chart:interval-major": 2.5,
        "chart:interval-minor-divisor": 2,
        "chart:japanese-candle-stick": false,
        "chart:label-arrangement": "side-by-side",
        "chart:label-position": "above",
        "chart:label-position-negative": "below",
        "chart:lines": true,
        "chart:link-data-style-to-source": true,
        "chart:logarithmic": false,
        "chart:maximum": 10.5,
        "chart:mean-value": true,
        "chart:minimum": 0.5,
        "chart:origin": 1,
        "chart:overlap": 20,
        "chart:percentage": false,
        "chart:pie-offset": 10,
        "chart:regression-force-intercept": true,
        "chart:regression-intercept-value": 1.25,
        "chart:regression-max-degree": 3,
        "chart:regression-moving-type": "prior",
        "chart:regression-name": "Trend",
        "chart:regression-period": 2,
        "chart:regression-type": "polynomial",
        "chart:reverse-direction": false,
        "chart:right-angled-axes": true,
        "chart:scale-text": false,
        "chart:series-source": "columns",
        "chart:solid-type": "cuboid",
        "chart:sort-by-x-values": true,
        "chart:spline-order": 3,
        "chart:spline-resolution": 5,
        "chart:stacked": false,
        "chart:symbol-height": 360000,
        "chart:symbol-type": "image",
        "chart:symbol-width": 360000,
        "chart:text-overlap": true,
        "chart:three-dimensional": false,
        "chart:tick-mark-position": "atLabels",
        "chart:tick-marks-major-inner": true,
        "chart:tick-marks-major-outer": false,
        "chart:tick-marks-minor-inner": true,
        "chart:tick-marks-minor-outer": false,
        "chart:treat-empty-cells": "leave-gap",
        "chart:vertical": true,
        "chart:visible": true,
        "style:direction": "ltr",
        "style:rotation-angle": 90,
        "text:line-break": true,
      },
      symbolImage: { attributes: { "xlink:href": "media/symbol.png" } },
    };
    expect(parseChartStyleProperties(element(chartPropertiesXml(options, "emu")), "emu")).toEqual({
      ok: true,
      value: options,
    });
    const namedSymbol: ChartStylePropertiesOptions = {
      name: "style:chart-properties",
      attributes: { "chart:symbol-name": "square", "chart:symbol-type": "named-symbol" },
    };
    expect(
      parseChartStyleProperties(element(chartPropertiesXml(namedSymbol, "emu")), "emu"),
    ).toEqual({ ok: true, value: namedSymbol });
  });

  it("round-trips list level properties and label alignment", () => {
    const options: ListLevelStylePropertiesOptions = {
      name: "style:list-level-properties",
      attributes: {
        "fo:width": 360000,
        "fo:height": 360000,
        "fo:text-align": "start",
        "style:font-name": "OpenSymbol",
        "style:vertical-pos": "top",
        "style:vertical-rel": "baseline",
        "svg:y": 0,
        "text:list-level-position-and-space-mode": "label-alignment",
        "text:min-label-distance": 72000,
        "text:min-label-width": 288000,
        "text:space-before": 36000,
      },
      labelAlignment: {
        attributes: {
          "text:label-followed-by": "listtab",
          "text:list-tab-stop-position": 432000,
          "fo:text-indent": -108000,
          "fo:margin-left": 360000,
        },
      },
    };
    expect(
      parseListLevelProperties(element(listLevelPropertiesXml(options, "emu")), "emu"),
    ).toEqual({ ok: true, value: options });
  });

  it("requires label alignment to identify its following content", () => {
    expect(
      parseListLevelProperties(
        element(
          '<style:list-level-properties><style:list-level-label-alignment fo:margin-left="1cm"/></style:list-level-properties>',
        ),
        "hundredth-millimeter",
      ),
    ).toMatchObject({
      ok: false,
      attributeName: "text:label-followed-by",
      reason: "required style property attribute is missing",
    });
  });

  it("round-trips linked and embedded background images", () => {
    const linked =
      '<style:paragraph-properties><style:background-image xlink:href="https://example.com/a.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" style:repeat="repeat" style:position="left top" style:filter-name="" draw:opacity="80%"/></style:paragraph-properties>';
    expect(stylePropertyXml(property(linked), "emu")).toBe(linked);

    const embedded = property(
      '<style:paragraph-properties><style:background-image xlink:href="#payload" style:repeat="stretch"><office:binary-data>YWJj</office:binary-data></style:background-image></style:paragraph-properties>',
    );
    expect(stylePropertyXml(embedded, "emu")).toBe(
      '<style:paragraph-properties><style:background-image xlink:href="#payload" style:repeat="stretch"><office:binary-data>YWJj</office:binary-data></style:background-image></style:paragraph-properties>',
    );
  });

  it("round-trips every remaining style property child surface", () => {
    const sources = [
      '<style:paragraph-properties><style:tab-stops><style:tab-stop style:position="2cm" style:type="char" style:char="," style:leader-type="solid" style:leader-style="dotted" style:leader-width="0.05cm" style:leader-color="#ff0000" style:leader-text="-" style:leader-text-style="T1"/></style:tab-stops></style:paragraph-properties>',
      '<style:paragraph-properties><style:drop-cap style:length="word" style:lines="3" style:distance="0.2cm" style:style-name="DropCap"/></style:paragraph-properties>',
      '<style:graphic-properties><style:columns fo:column-count="2" fo:column-gap="0.5cm"><style:column-sep style:style="dashed" style:width="0.05cm" style:height="80%" style:vertical-align="top" style:color="#123456"/><style:column fo:start-indent="0.1cm" fo:end-indent="0.2cm" fo:space-before="0.3cm" fo:space-after="0.4cm" style:rel-width="1*"/></style:columns></style:graphic-properties>',
      '<style:page-layout-properties><style:footnote-sep style:width="0.05cm" style:rel-width="25%" style:color="#123456" style:line-style="solid" style:adjustment="center" style:distance-before-sep="0.1cm" style:distance-after-sep="0.2cm"/></style:page-layout-properties>',
      '<style:drawing-page-properties><presentation:sound xlink:href="media/sound.wav" xlink:type="simple" xlink:show="replace" xlink:actuate="onRequest" presentation:play-full="true" xml:id="sound"/></style:drawing-page-properties>',
      '<style:section-properties><text:notes-configuration text:note-class="footnote" text:citation-style-name="Cite" text:citation-body-style-name="CiteBody" text:default-style-name="Note" text:master-page-name="NotePage" text:start-value="1" style:num-format="1" style:num-prefix="[" style:num-suffix="]" text:start-numbering-at="document" text:footnotes-position="page"><text:note-continuation-notice-forward>continued</text:note-continuation-notice-forward><text:note-continuation-notice-backward>end</text:note-continuation-notice-backward></text:notes-configuration></style:section-properties>',
      '<style:graphic-properties><text:list-style style:name="GraphicList" style:display-name="Graphic List" text:consecutive-numbering="true"><text:list-level-style-number text:level="1" style:num-format="1" style:num-prefix="(" style:num-suffix=")" text:display-levels="1" text:start-value="1" text:style-name="NumberStyle"><style:list-level-properties text:list-level-position-and-space-mode="label-width-and-position" text:min-label-width="0.8cm"/><style:text-properties fo:font-weight="bold"/></text:list-level-style-number><text:list-level-style-bullet text:level="2" text:bullet-char="•" text:bullet-relative-size="80%" text:style-name="BulletStyle"><style:list-level-properties text:space-before="0.2cm"/></text:list-level-style-bullet><text:list-level-style-image text:level="3" xlink:href="media/bullet.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><style:list-level-properties fo:width="0.4cm" fo:height="0.4cm"/></text:list-level-style-image></text:list-style></style:graphic-properties>',
    ];
    for (const source of sources) {
      const family = childFamily(source);
      const parsed = propertyIn(source, family);
      expect(stylePropertyXml(parsed, "emu")).toBe(source);
    }

    function childFamily(source: string) {
      if (source.startsWith("<style:paragraph-properties")) return "paragraph" as const;
      if (source.startsWith("<style:graphic-properties")) return "graphic" as const;
      if (source.startsWith("<style:page-layout-properties")) return "page-layout" as const;
      if (source.startsWith("<style:drawing-page-properties")) return "drawing-page" as const;
      if (source.startsWith("<style:section-properties")) return "section" as const;
      throw new Error(`Unsupported child fixture: ${source}`);
    }
  });

  it("rejects unknown rebuilt style content structurally", () => {
    expect(
      parseStyleProperty(element('<style:paragraph-properties officeooo:unknown="1"/>'), {
        family: "paragraph",
        scale: "emu",
      }),
    ).toMatchObject({ ok: false, attributeName: "officeooo:unknown" });
    expect(
      parseStyleProperty(
        element("<style:paragraph-properties><unknown/></style:paragraph-properties>"),
        {
          family: "paragraph",
          scale: "emu",
        },
      ),
    ).toMatchObject({ ok: false, elementName: "unknown" });
    expect(
      parseChartStyleProperties(
        element("<style:chart-properties><unknown/></style:chart-properties>"),
        "emu",
      ),
    ).toMatchObject({ ok: false, elementName: "unknown" });
    expect(
      parseChartStyleProperties(
        element(
          '<style:chart-properties><chart:symbol-image xlink:href="a"/><chart:symbol-image xlink:href="b"/></style:chart-properties>',
        ),
        "emu",
      ),
    ).toMatchObject({ ok: false, reason: "duplicate style property child" });
    expect(
      parseListLevelProperties(
        element("<style:list-level-properties><unknown/></style:list-level-properties>"),
        "emu",
      ),
    ).toMatchObject({ ok: false, elementName: "unknown" });
  });
});
