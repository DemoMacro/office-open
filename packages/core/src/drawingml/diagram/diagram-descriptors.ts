/**
 * Diagram descriptors for declarative XML read/write.
 *
 * @module
 */

import { escapeXml, findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import type { DiagramExtensionListOptions } from "./diagram-props";
import type { DiagramRelationshipIdsOptions } from "./diagram-rel";
import type { DiagramStyleOptions } from "./diagram-style";
import type { PresentationLayoutVariablesOptions } from "./layout-vars";

// ── dgm:relIds descriptor ──

export const diagramRelationshipIdsDesc: CustomDescriptor<DiagramRelationshipIdsOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return `<dgm:relIds r:dm="${escapeXml(opts.dm)}" r:lo="${escapeXml(opts.lo)}" r:qs="${escapeXml(opts.qs)}" r:cs="${escapeXml(opts.cs)}"/>`;
  },
  parse(el, _ctx) {
    const result: Partial<DiagramRelationshipIdsOptions> = {};
    if (el.attributes) {
      if (el.attributes["r:dm"] !== undefined) result.dm = String(el.attributes["r:dm"]);
      if (el.attributes["r:lo"] !== undefined) result.lo = String(el.attributes["r:lo"]);
      if (el.attributes["r:qs"] !== undefined) result.qs = String(el.attributes["r:qs"]);
      if (el.attributes["r:cs"] !== undefined) result.cs = String(el.attributes["r:cs"]);
    }
    return result as DiagramRelationshipIdsOptions;
  },
};

// ── dgm:style descriptor ──

export const diagramStyleDesc: CustomDescriptor<DiagramStyleOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const lineIdx = opts.lineReference?.idx ?? 1;
    const fillIdx = opts.fillReference?.idx ?? 1;
    const effectIdx = opts.effectReference?.idx ?? 0;
    const fontIdx = opts.fontReference?.idx ?? "minor";

    return (
      `<dgm:style>` +
      `<a:lnRef idx="${lineIdx}"><a:schemeClr val="accent1"/></a:lnRef>` +
      `<a:fillRef idx="${fillIdx}"><a:schemeClr val="accent1"/></a:fillRef>` +
      `<a:effectRef idx="${effectIdx}"><a:schemeClr val="accent1"/></a:effectRef>` +
      `<a:fontRef idx="${escapeXml(fontIdx)}"><a:schemeClr val="tx1"/></a:fontRef>` +
      `</dgm:style>`
    );
  },
  parse(el, _ctx) {
    const result: Partial<DiagramStyleOptions> = {};

    const lnRef = findChild(el, "a:lnRef");
    if (lnRef?.attributes?.["idx"] !== undefined)
      result.lineReference = { idx: Number(lnRef.attributes["idx"]) };

    const fillRef = findChild(el, "a:fillRef");
    if (fillRef?.attributes?.["idx"] !== undefined)
      result.fillReference = { idx: Number(fillRef.attributes["idx"]) };

    const effectRef = findChild(el, "a:effectRef");
    if (effectRef?.attributes?.["idx"] !== undefined)
      result.effectReference = { idx: Number(effectRef.attributes["idx"]) };

    const fontRef = findChild(el, "a:fontRef");
    if (fontRef?.attributes?.["idx"] !== undefined)
      result.fontReference = { idx: String(fontRef.attributes["idx"]) };

    return result as DiagramStyleOptions;
  },
};

// ── dgm:presLayoutVars descriptor ──

export const presentationLayoutVariablesDesc: CustomDescriptor<PresentationLayoutVariablesOptions> =
  {
    kind: "custom",
    stringify(opts, _ctx) {
      const parts: string[] = [];

      if (opts.orgChart?.val !== undefined)
        parts.push(`<dgm:orgChart val="${opts.orgChart.val ? 1 : 0}"/>`);

      if (opts.maxChildren?.val !== undefined)
        parts.push(`<dgm:chMax val="${opts.maxChildren.val}"/>`);

      if (opts.preferredChildren?.val !== undefined)
        parts.push(`<dgm:chPref val="${opts.preferredChildren.val}"/>`);

      if (opts.animateOneByOne?.val !== undefined)
        parts.push(`<dgm:animOne val="${escapeXml(opts.animateOneByOne.val)}"/>`);

      if (opts.animationLevel?.val !== undefined)
        parts.push(`<dgm:animLvl val="${escapeXml(opts.animationLevel.val)}"/>`);

      if (opts.hierBranch?.val !== undefined)
        parts.push(`<dgm:hierBranch val="${escapeXml(opts.hierBranch.val)}"/>`);

      if (parts.length === 0) return `<dgm:presLayoutVars/>`;
      return `<dgm:presLayoutVars>${parts.join("")}</dgm:presLayoutVars>`;
    },
    parse(el, _ctx) {
      const result: Partial<PresentationLayoutVariablesOptions> = {};

      const orgChart = findChild(el, "dgm:orgChart");
      if (orgChart?.attributes?.["val"] !== undefined)
        result.orgChart = {
          val: orgChart.attributes["val"] === 1 || orgChart.attributes["val"] === "1",
        };

      const chMax = findChild(el, "dgm:chMax");
      if (chMax?.attributes?.["val"] !== undefined)
        result.maxChildren = { val: Number(chMax.attributes["val"]) };

      const chPref = findChild(el, "dgm:chPref");
      if (chPref?.attributes?.["val"] !== undefined)
        result.preferredChildren = { val: Number(chPref.attributes["val"]) };

      const animOne = findChild(el, "dgm:animOne");
      if (animOne?.attributes?.["val"] !== undefined)
        result.animateOneByOne = {
          val: String(
            animOne.attributes["val"],
          ) as PresentationLayoutVariablesOptions["animateOneByOne"] extends
            | { val?: infer V }
            | undefined
            ? V
            : never,
        };

      const animLvl = findChild(el, "dgm:animLvl");
      if (animLvl?.attributes?.["val"] !== undefined)
        result.animationLevel = {
          val: String(
            animLvl.attributes["val"],
          ) as PresentationLayoutVariablesOptions["animationLevel"] extends
            | { val?: infer V }
            | undefined
            ? V
            : never,
        };

      const hierBranch = findChild(el, "dgm:hierBranch");
      if (hierBranch?.attributes?.["val"] !== undefined)
        result.hierBranch = {
          val: String(
            hierBranch.attributes["val"],
          ) as PresentationLayoutVariablesOptions["hierBranch"] extends
            | { val?: infer V }
            | undefined
            ? V
            : never,
        };

      return result as PresentationLayoutVariablesOptions;
    },
  };

// ── dgm:extLst descriptor ──

export const diagramExtensionListDesc: CustomDescriptor<DiagramExtensionListOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    if (!opts.extensions?.length) return undefined;

    const inner = opts.extensions.map((ext) => `<a:ext uri="${escapeXml(ext.uri)}"/>`).join("");

    return `<dgm:extLst>${inner}</dgm:extLst>`;
  },
  parse(el, _ctx) {
    const result: Partial<DiagramExtensionListOptions> = {};

    if (el.elements) {
      const extensions = el.elements
        .filter((c) => c.name === "a:ext")
        .map((c) => ({ uri: String(c.attributes?.["uri"] ?? "") }));
      if (extensions.length) result.extensions = extensions;
    }

    return result as DiagramExtensionListOptions;
  },
};
