/**
 * Volatile function types — the standalone `xl/volTypes.xml` part
 * (CT_VolTypes root, RTD server topic registrations). Never a child of
 * workbook.xml: sml.xsd declares volTypes as a top-level part element.
 *
 * @module
 */
import { attr, attrNum, escapeXml, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type {
  VolMainOptions,
  VolTopicOptions,
  VolTopicRefOptions,
  VolTypeOptions,
} from "./workbook/types";

export type { VolTypeOptions } from "./workbook/types";

const SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** Emit the complete `<volTypes>` root element for xl/volTypes.xml. */
export function buildVolTypesXml(volTypes: readonly VolTypeOptions[]): string {
  if (volTypes.length === 0) return "";
  const parts: string[] = [`<volTypes xmlns="${SML_NS}" count="${volTypes.length}">`];
  for (const vt of volTypes) {
    const vtType = vt.type ?? "realTimeData";
    const mains = vt.mains ?? [];
    if (mains.length > 0) {
      const mainParts: string[] = [];
      for (const m of mains) {
        const tpParts: string[] = [];
        for (const topic of m.topics ?? []) {
          const tpInner: string[] = [`<v>${escapeXml(topic.value)}</v>`];
          for (const stp of topic.stringTopics ?? []) {
            tpInner.push(`<stp>${escapeXml(stp)}</stp>`);
          }
          for (const tr of topic.refs ?? []) {
            tpInner.push(`<tr r="${escapeXml(tr.reference)}" s="${tr.sheetIndex}"/>`);
          }
          const tpAttr =
            topic.valueType && topic.valueType !== "n" ? ` t="${escapeXml(topic.valueType)}"` : "";
          tpParts.push(`<tp${tpAttr}>${tpInner.join("")}</tp>`);
        }
        mainParts.push(`<main first="${escapeXml(m.first)}">${tpParts.join("")}</main>`);
      }
      parts.push(`<volType type="${vtType}">${mainParts.join("")}</volType>`);
    } else {
      parts.push(`<volType type="${vtType}"/>`);
    }
  }
  parts.push("</volTypes>");
  return parts.join("");
}

/** Parse the `<volTypes>` root element of xl/volTypes.xml. */
export function parseVolTypesEl(el: Element): VolTypeOptions[] {
  const volTypes: VolTypeOptions[] = [];
  for (const vt of el.elements ?? []) {
    if (vt.name !== "volType") continue;
    const volType: VolTypeOptions = {};
    const typeVal = attr(vt, "type");
    if (typeVal) volType.type = typeVal as VolTypeOptions["type"];
    const mains: VolMainOptions[] = [];
    for (const m of vt.elements ?? []) {
      if (m.name !== "main") continue;
      const main: VolMainOptions = { first: attr(m, "first") ?? "" };
      const topics: VolTopicOptions[] = [];
      for (const tp of m.elements ?? []) {
        if (tp.name !== "tp") continue;
        const vEl = findChild(tp, "v");
        const topic: VolTopicOptions = { value: String(vEl?.elements?.[0]?.text ?? "") };
        const tVal = attr(tp, "t");
        if (tVal) topic.valueType = tVal;
        const stps: string[] = [];
        const refs: VolTopicRefOptions[] = [];
        for (const inner of tp.elements ?? []) {
          if (inner.name === "stp") stps.push(String(inner.elements?.[0]?.text ?? ""));
          if (inner.name === "tr") {
            const ref: VolTopicRefOptions = {
              reference: attr(inner, "r") ?? "",
              sheetIndex: attrNum(inner, "s") ?? 0,
            };
            refs.push(ref);
          }
        }
        if (stps.length > 0) topic.stringTopics = stps;
        if (refs.length > 0) topic.refs = refs;
        topics.push(topic);
      }
      if (topics.length > 0) main.topics = topics;
      mains.push(main);
    }
    if (mains.length > 0) volType.mains = mains;
    volTypes.push(volType);
  }
  return volTypes;
}
