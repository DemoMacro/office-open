import type { PropertiesOptions } from "@file/core-properties";
import { attr, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/**
 * Parse word/settings.xml Element into PropertiesOptions fields.
 */
export function parseSettings(el: Element | undefined): Partial<PropertiesOptions> {
  if (!el) return {};

  const opts: Record<string, unknown> = {};

  // evenAndOddHeaderAndFooters → w:evenAndOddHeaders (check val attribute)
  const eohEl = findChild(el, "w:evenAndOddHeaders");
  if (eohEl) {
    const val = attr(eohEl, "w:val");
    opts.evenAndOddHeaderAndFooters = val !== "false" && val !== "0" && val !== "off";
  }

  // view → w:view/@w:val
  const viewEl = findChild(el, "w:view");
  if (viewEl) {
    const val = attr(viewEl, "w:val");
    if (val) opts.view = val as PropertiesOptions["view"];
  }

  // zoom → w:zoom/@w:percent, @w:val
  const zoomEl = findChild(el, "w:zoom");
  if (zoomEl) {
    const zoom: Record<string, unknown> = {};
    const percent = attr(zoomEl, "w:percent");
    if (percent) zoom.percent = parseInt(percent, 10);
    const val = attr(zoomEl, "w:val");
    if (val) zoom.val = val;
    if (Object.keys(zoom).length > 0) opts.zoom = zoom;
  }

  // defaultTabStop → w:defaultTabStop/@w:val
  const tabStopEl = findChild(el, "w:defaultTabStop");
  if (tabStopEl) {
    const val = attr(tabStopEl, "w:val");
    if (val) opts.defaultTabStop = parseInt(val, 10);
  }

  // features.trackRevisions → w:trackRevisions (presence)
  if (findChild(el, "w:trackRevisions")) {
    const features = (opts.features as Record<string, unknown>) ?? {};
    (features as Record<string, unknown>).trackRevisions = true;
    opts.features = features;
  }

  // features.updateFields → w:updateFields (presence)
  if (findChild(el, "w:updateFields")) {
    const features = (opts.features as Record<string, unknown>) ?? {};
    (features as Record<string, unknown>).updateFields = true;
    opts.features = features;
  }

  // compatabilityModeVersion → w:compat/w:compatSetting
  const compatEl = findChild(el, "w:compat");
  if (compatEl) {
    for (const child of compatEl.elements ?? []) {
      if (child.name === "w:compatSetting") {
        const name = attr(child, "w:name");
        if (name === "compatibilityMode") {
          const val = attr(child, "w:val");
          if (val) opts.compatabilityModeVersion = parseInt(val, 10);
        }
      }
    }
  }

  // docVars → w:docVars/w:docVar
  const docVarsEl = findChild(el, "w:docVars");
  if (docVarsEl) {
    const vars: Array<{ name: string; val: string }> = [];
    for (const child of docVarsEl.elements ?? []) {
      if (child.name === "w:docVar") {
        const name = attr(child, "w:name");
        const val = attr(child, "w:val");
        if (name && val) vars.push({ name, val });
      }
    }
    if (vars.length > 0) opts.docVars = vars;
  }

  // characterSpacingControl → w:characterSpacingControl/@w:val
  const cscEl = findChild(el, "w:characterSpacingControl");
  if (cscEl) {
    const val = attr(cscEl, "w:val");
    if (val) opts.characterSpacingControl = val as PropertiesOptions["characterSpacingControl"];
  }

  // displayBackgroundShape → w:displayBackgroundShape (presence)
  if (findChild(el, "w:displayBackgroundShape")) {
    opts.displayBackgroundShape = true;
  }

  // embedTrueTypeFonts → w:embedTrueTypeFonts (presence)
  if (findChild(el, "w:embedTrueTypeFonts")) {
    opts.embedTrueTypeFonts = true;
  }

  // embedSystemFonts → w:embedSystemFonts (presence)
  if (findChild(el, "w:embedSystemFonts")) {
    opts.embedSystemFonts = true;
  }

  // saveSubsetFonts → w:saveSubsetFonts (presence)
  if (findChild(el, "w:saveSubsetFonts")) {
    opts.saveSubsetFonts = true;
  }

  return opts;
}
