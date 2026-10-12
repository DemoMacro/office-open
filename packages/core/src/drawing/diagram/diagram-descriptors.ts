/**
 * Diagram descriptors for declarative XML read/write.
 *
 * @module
 */

import { escapeXml, findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { parseShapeStyle, stringifyShapeStyle } from "../../theme/style-matrix";
import type { DefaultShapeStyleOptions } from "../../theme/theme-options";
import {
  xsdAnimLevel,
  xsdDiagramDirection,
  xsdHierBranch,
  xsdResizeHandles,
} from "../../util/mappings";
import { parseOnOff } from "../../util/values";
import { parseDiagramExtensions, stringifyDiagramExtensions } from "./diagram-props";
import type { DiagramExtensionListOptions } from "./diagram-props";
import type { DiagramRelationshipIdsOptions } from "./diagram-rel";
import type { PresentationLayoutVariablesOptions } from "./layout-vars";

// ── dgm:relIds descriptor ──

export const diagramRelationshipIdsDesc: CustomDescriptor<DiagramRelationshipIdsOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return `<dgm:relIds r:dm="${escapeXml(opts.dataModel)}" r:lo="${escapeXml(opts.layout)}" r:qs="${escapeXml(opts.quickStyle)}" r:cs="${escapeXml(opts.colorStyle)}"/>`;
  },
  parse(el, _ctx) {
    const result: Partial<DiagramRelationshipIdsOptions> = {};
    if (el.attributes) {
      if (el.attributes["r:dm"] !== undefined) result.dataModel = String(el.attributes["r:dm"]);
      if (el.attributes["r:lo"] !== undefined) result.layout = String(el.attributes["r:lo"]);
      if (el.attributes["r:qs"] !== undefined) result.quickStyle = String(el.attributes["r:qs"]);
      if (el.attributes["r:cs"] !== undefined) result.colorStyle = String(el.attributes["r:cs"]);
    }
    return result as DiagramRelationshipIdsOptions;
  },
};

// ── dgm:style descriptor ──

export const diagramStyleDesc: CustomDescriptor<DefaultShapeStyleOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return stringifyShapeStyle(opts, _ctx, "dgm:style");
  },
  parse(el, _ctx) {
    return parseShapeStyle(el, _ctx) ?? ({} as DefaultShapeStyleOptions);
  },
};

// ── dgm:presLayoutVars descriptor ──

export const presentationLayoutVariablesDesc: CustomDescriptor<PresentationLayoutVariablesOptions> =
  {
    kind: "custom",
    stringify(opts, _ctx) {
      const parts: string[] = [];

      if (opts.orgChart)
        parts.push(
          opts.orgChart.val === undefined
            ? "<dgm:orgChart/>"
            : `<dgm:orgChart val="${opts.orgChart.val ? 1 : 0}"/>`,
        );

      if (opts.maxChildren)
        parts.push(
          opts.maxChildren.val === undefined
            ? "<dgm:chMax/>"
            : `<dgm:chMax val="${opts.maxChildren.val}"/>`,
        );

      if (opts.preferredChildren)
        parts.push(
          opts.preferredChildren.val === undefined
            ? "<dgm:chPref/>"
            : `<dgm:chPref val="${opts.preferredChildren.val}"/>`,
        );

      if (opts.bulletEnabled)
        parts.push(
          opts.bulletEnabled.val === undefined
            ? "<dgm:bulletEnabled/>"
            : `<dgm:bulletEnabled val="${opts.bulletEnabled.val ? 1 : 0}"/>`,
        );

      if (opts.direction)
        parts.push(
          opts.direction.val === undefined
            ? "<dgm:dir/>"
            : `<dgm:dir val="${xsdDiagramDirection.to(opts.direction.val)}"/>`,
        );

      if (opts.animateOneByOne)
        parts.push(
          opts.animateOneByOne.val === undefined
            ? "<dgm:animOne/>"
            : `<dgm:animOne val="${escapeXml(opts.animateOneByOne.val)}"/>`,
        );

      if (opts.animationLevel)
        parts.push(
          opts.animationLevel.val === undefined
            ? "<dgm:animLvl/>"
            : `<dgm:animLvl val="${xsdAnimLevel.to(opts.animationLevel.val)}"/>`,
        );

      if (opts.hierBranch)
        parts.push(
          opts.hierBranch.val === undefined
            ? "<dgm:hierBranch/>"
            : `<dgm:hierBranch val="${xsdHierBranch.to(opts.hierBranch.val)}"/>`,
        );

      if (opts.resizeHandles)
        parts.push(
          opts.resizeHandles.val === undefined
            ? "<dgm:resizeHandles/>"
            : `<dgm:resizeHandles val="${xsdResizeHandles.to(opts.resizeHandles.val)}"/>`,
        );

      if (parts.length === 0) return `<dgm:presLayoutVars/>`;
      return `<dgm:presLayoutVars>${parts.join("")}</dgm:presLayoutVars>`;
    },
    parse(el, _ctx) {
      const result: PresentationLayoutVariablesOptions = {};

      const orgChart = findChild(el, "dgm:orgChart");
      if (orgChart)
        result.orgChart =
          orgChart.attributes?.["val"] === undefined
            ? {}
            : { val: parseOnOff(orgChart.attributes["val"]) ?? false };

      const chMax = findChild(el, "dgm:chMax");
      if (chMax)
        result.maxChildren =
          chMax.attributes?.["val"] === undefined ? {} : { val: Number(chMax.attributes["val"]) };

      const chPref = findChild(el, "dgm:chPref");
      if (chPref)
        result.preferredChildren =
          chPref.attributes?.["val"] === undefined ? {} : { val: Number(chPref.attributes["val"]) };

      const bulletEnabled = findChild(el, "dgm:bulletEnabled");
      if (bulletEnabled)
        result.bulletEnabled =
          bulletEnabled.attributes?.["val"] === undefined
            ? {}
            : { val: parseOnOff(bulletEnabled.attributes["val"]) ?? false };

      const direction = findChild(el, "dgm:dir");
      if (direction)
        result.direction =
          direction.attributes?.["val"] === undefined
            ? {}
            : {
                val: xsdDiagramDirection.from(String(direction.attributes["val"])) as
                  | "normal"
                  | "reversed"
                  | undefined,
              };

      const animOne = findChild(el, "dgm:animOne");
      if (animOne)
        result.animateOneByOne =
          animOne.attributes?.["val"] === undefined
            ? {}
            : {
                val: String(
                  animOne.attributes["val"],
                ) as PresentationLayoutVariablesOptions["animateOneByOne"] extends
                  | { val?: infer V }
                  | undefined
                  ? V
                  : never,
              };

      const animLvl = findChild(el, "dgm:animLvl");
      if (animLvl)
        result.animationLevel =
          animLvl.attributes?.["val"] === undefined
            ? {}
            : {
                val: xsdAnimLevel.from(
                  String(animLvl.attributes["val"]),
                ) as PresentationLayoutVariablesOptions["animationLevel"] extends
                  | { val?: infer V }
                  | undefined
                  ? V
                  : never,
              };

      const hierBranch = findChild(el, "dgm:hierBranch");
      if (hierBranch)
        result.hierBranch =
          hierBranch.attributes?.["val"] === undefined
            ? {}
            : {
                val: xsdHierBranch.from(
                  String(hierBranch.attributes["val"]),
                ) as PresentationLayoutVariablesOptions["hierBranch"] extends
                  | { val?: infer V }
                  | undefined
                  ? V
                  : never,
              };

      const resizeHandles = findChild(el, "dgm:resizeHandles");
      if (resizeHandles)
        result.resizeHandles =
          resizeHandles.attributes?.["val"] === undefined
            ? {}
            : {
                val: xsdResizeHandles.from(String(resizeHandles.attributes["val"])) as
                  | "exact"
                  | "relative"
                  | undefined,
              };

      return result;
    },
  };

// ── dgm:extLst descriptor ──

export const diagramExtensionListDesc: CustomDescriptor<DiagramExtensionListOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return opts.extensions?.length ? stringifyDiagramExtensions(opts.extensions) : undefined;
  },
  parse(el, _ctx) {
    const result: DiagramExtensionListOptions = {};

    const extensions = parseDiagramExtensions(el);
    if (extensions.length) {
      result.extensions = extensions;
    }

    return result;
  },
};
