/**
 * Diagram descriptors for declarative XML read/write.
 *
 * @module
 */

import { escapeXml, findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import type { DiagramExtLstOptions } from "./diagram-props";
import type { DiagramRelIdsOptions } from "./diagram-rel";
import type { DiagramStyleOptions } from "./diagram-style";
import type { PresLayoutVarsOptions } from "./layout-vars";

// ── dgm:relIds descriptor ──

export const diagramRelIdsDesc: CustomDescriptor<DiagramRelIdsOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return `<dgm:relIds r:dm="${escapeXml(opts.dm)}" r:lo="${escapeXml(opts.lo)}" r:qs="${escapeXml(opts.qs)}" r:cs="${escapeXml(opts.cs)}"/>`;
  },
  parse(el, _ctx) {
    const result: Partial<DiagramRelIdsOptions> = {};
    if (el.attributes) {
      if (el.attributes["r:dm"] !== undefined) result.dm = String(el.attributes["r:dm"]);
      if (el.attributes["r:lo"] !== undefined) result.lo = String(el.attributes["r:lo"]);
      if (el.attributes["r:qs"] !== undefined) result.qs = String(el.attributes["r:qs"]);
      if (el.attributes["r:cs"] !== undefined) result.cs = String(el.attributes["r:cs"]);
    }
    return result;
  },
};

// ── dgm:style descriptor ──

export const diagramStyleDesc: CustomDescriptor<DiagramStyleOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const lnIdx = opts.lnIdx ?? 1;
    const fillIdx = opts.fillIdx ?? 1;
    const effectIdx = opts.effectIdx ?? 0;
    const fontIdx = opts.fontIdx ?? "minor";

    return (
      `<dgm:style>` +
      `<a:lnRef idx="${lnIdx}"><a:schemeClr val="accent1"/></a:lnRef>` +
      `<a:fillRef idx="${fillIdx}"><a:schemeClr val="accent1"/></a:fillRef>` +
      `<a:effectRef idx="${effectIdx}"><a:schemeClr val="accent1"/></a:effectRef>` +
      `<a:fontRef idx="${escapeXml(fontIdx)}"><a:schemeClr val="tx1"/></a:fontRef>` +
      `</dgm:style>`
    );
  },
  parse(el, _ctx) {
    const result: Partial<DiagramStyleOptions> = {};

    const lnRef = findChild(el, "a:lnRef");
    if (lnRef?.attributes?.["idx"] !== undefined) result.lnIdx = Number(lnRef.attributes["idx"]);

    const fillRef = findChild(el, "a:fillRef");
    if (fillRef?.attributes?.["idx"] !== undefined)
      result.fillIdx = Number(fillRef.attributes["idx"]);

    const effectRef = findChild(el, "a:effectRef");
    if (effectRef?.attributes?.["idx"] !== undefined)
      result.effectIdx = Number(effectRef.attributes["idx"]);

    const fontRef = findChild(el, "a:fontRef");
    if (fontRef?.attributes?.["idx"] !== undefined)
      result.fontIdx = String(fontRef.attributes["idx"]);

    return result;
  },
};

// ── dgm:presLayoutVars descriptor ──

export const presLayoutVarsDesc: CustomDescriptor<PresLayoutVarsOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const parts: string[] = [];

    if (opts.orgChart?.val !== undefined)
      parts.push(`<dgm:orgChart val="${opts.orgChart.val ? 1 : 0}"/>`);

    if (opts.chMax?.val !== undefined) parts.push(`<dgm:chMax val="${opts.chMax.val}"/>`);

    if (opts.chPref?.val !== undefined) parts.push(`<dgm:chPref val="${opts.chPref.val}"/>`);

    if (opts.animOne?.val !== undefined)
      parts.push(`<dgm:animOne val="${escapeXml(opts.animOne.val)}"/>`);

    if (opts.animLvl?.val !== undefined)
      parts.push(`<dgm:animLvl val="${escapeXml(opts.animLvl.val)}"/>`);

    if (opts.hierBranch?.val !== undefined)
      parts.push(`<dgm:hierBranch val="${escapeXml(opts.hierBranch.val)}"/>`);

    if (parts.length === 0) return `<dgm:presLayoutVars/>`;
    return `<dgm:presLayoutVars>${parts.join("")}</dgm:presLayoutVars>`;
  },
  parse(el, _ctx) {
    const result: Partial<PresLayoutVarsOptions> = {};

    const orgChart = findChild(el, "dgm:orgChart");
    if (orgChart?.attributes?.["val"] !== undefined)
      result.orgChart = {
        val: orgChart.attributes["val"] === 1 || orgChart.attributes["val"] === "1",
      };

    const chMax = findChild(el, "dgm:chMax");
    if (chMax?.attributes?.["val"] !== undefined)
      result.chMax = { val: Number(chMax.attributes["val"]) };

    const chPref = findChild(el, "dgm:chPref");
    if (chPref?.attributes?.["val"] !== undefined)
      result.chPref = { val: Number(chPref.attributes["val"]) };

    const animOne = findChild(el, "dgm:animOne");
    if (animOne?.attributes?.["val"] !== undefined)
      result.animOne = {
        val: String(animOne.attributes["val"]) as PresLayoutVarsOptions["animOne"] extends
          | { val?: infer V }
          | undefined
          ? V
          : never,
      };

    const animLvl = findChild(el, "dgm:animLvl");
    if (animLvl?.attributes?.["val"] !== undefined)
      result.animLvl = {
        val: String(animLvl.attributes["val"]) as PresLayoutVarsOptions["animLvl"] extends
          | { val?: infer V }
          | undefined
          ? V
          : never,
      };

    const hierBranch = findChild(el, "dgm:hierBranch");
    if (hierBranch?.attributes?.["val"] !== undefined)
      result.hierBranch = {
        val: String(hierBranch.attributes["val"]) as PresLayoutVarsOptions["hierBranch"] extends
          | { val?: infer V }
          | undefined
          ? V
          : never,
      };

    return result;
  },
};

// ── dgm:extLst descriptor ──

export const diagramExtLstDesc: CustomDescriptor<DiagramExtLstOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    if (!opts.extensions?.length) return undefined;

    const inner = opts.extensions.map((ext) => `<a:ext uri="${escapeXml(ext.uri)}"/>`).join("");

    return `<dgm:extLst>${inner}</dgm:extLst>`;
  },
  parse(el, _ctx) {
    const result: Partial<DiagramExtLstOptions> = {};

    if (el.elements) {
      const extensions = el.elements
        .filter((c) => c.name === "a:ext")
        .map((c) => ({ uri: String(c.attributes?.["uri"] ?? "") }));
      if (extensions.length) result.extensions = extensions;
    }

    return result;
  },
};
