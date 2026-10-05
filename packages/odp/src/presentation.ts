import type { FormContainerOptions, FormControlOptions } from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  CHART_MIME,
  chartBodyXml,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  generateOcf,
  lengthToEmu,
  metaXml,
  OdfSchemaError,
  parseEmbeddedCharts,
  parseGraphicStyles,
  parseMeta,
  readOcf,
  readXml,
  textOf,
  xmlElement,
  type OdfPackageFiles,
} from "@office-open/odf";
import type { PresentationOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import type { OdpChart, OdpImage } from "./drawing";
import { OdpParseError } from "./error";
import { parseColumnWidths, parseSlide, slideXml } from "./slide";
import { parseTextStyles } from "./styles";

export const MIME = "application/vnd.oasis.opendocument.presentation";

export const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"',
  'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"',
  'xmlns:xforms="http://www.w3.org/2002/xforms"',
].join(" ");

export function generateOdp(options: PresentationOptions): Uint8Array {
  const styles: string[] = [];
  const images: OdpImage[] = [];
  const charts: OdpChart[] = [];
  const size = normalizeSize(options.size);
  const pageLayout = xmlElement("style:page-layout", { "style:name": "PM1" }, [
    xmlElement("style:page-layout-properties", {
      "fo:page-width": emuToLength(size.width),
      "fo:page-height": emuToLength(size.height),
    }),
  ]);
  const pages = (options.slides ?? []).map((slide, index) =>
    slideXml(slide, index + 1, styles, images, charts),
  );
  const files: OdfPackageFiles = {
    "content.xml": contentXml(pages.join(""), styles),
    "styles.xml": stylesXml(pageLayout),
    "meta.xml": metaXml(options),
  };
  for (const image of images) files[image.path] = image.data;
  for (const chart of charts) files[`${chart.path}/content.xml`] = chartBodyXml(chart.chart);
  return generateOcf(
    MIME,
    files,
    Object.fromEntries(charts.map((chart) => [`${chart.path}/`, CHART_MIME])),
  );
}

export function parseOdp(data: Uint8Array): PresentationOptions {
  try {
    return parseOdpPresentation(data);
  } catch (cause) {
    if (cause instanceof OdpParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdpParseError(cause.message, cause.part, cause.path, cause.name, cause.reason, {
        cause,
      });
    }
    throw new OdpParseError(
      cause instanceof Error ? cause.message : "Unable to parse ODP package",
      "mimetype",
      "/mimetype",
      "mimetype",
      "invalid ODP package",
      { cause },
    );
  }
}

export function parseOdpPresentation(data: Uint8Array): PresentationOptions {
  const { files, binaries, manifest } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:presentation");
  const stylesDocument = files["styles.xml"] ? readXml(files, "styles.xml") : undefined;
  const pageLayout = childNamed(
    childNamed(childNamed(stylesDocument, "office:automatic-styles"), "style:page-layout"),
    "style:page-layout-properties",
  );
  const width = lengthToEmu(attributeString(pageLayout, "fo:page-width"));
  const height = lengthToEmu(attributeString(pageLayout, "fo:page-height"));
  const graphicStyles = parseGraphicStyles(childNamed(content, "office:automatic-styles"));
  const chartPool = parseEmbeddedCharts(manifest, files);
  const result = {
    ...parseMeta(files),
    ...(width && height ? { size: { width, height } } : {}),
    slides: childrenNamed(body, "draw:page").map((page) =>
      parseSlide(
        page,
        parseTextStyles(childNamed(content, "office:automatic-styles")),
        parseColumnWidths(childNamed(content, "office:automatic-styles")),
        binaries,
        graphicStyles,
        chartPool,
      ),
    ),
  };
  return result;
}

export function contentXml(pages: string, styles: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
}

export function stylesXml(pageLayout: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/><office:automatic-styles>${pageLayout}</office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
}

export function normalizeSize(size: PresentationOptions["size"]): {
  width: number;
  height: number;
} {
  if (size === "4:3") return { width: 9144000, height: 6858000 };
  if (typeof size === "object" && size !== null) {
    return {
      width: typeof size.width === "number" ? size.width : (lengthToEmu(size.width) ?? 0),
      height: typeof size.height === "number" ? size.height : (lengthToEmu(size.height) ?? 0),
    };
  }
  return { width: 12192000, height: 6858000 };
}

export function formsXml(forms: FormContainerOptions[]): string {
  return xmlElement(
    "office:forms",
    {
      "form:automatic-focus": forms[0]?.automaticFocus,
      "form:apply-design-mode": forms[0]?.designMode,
    },
    forms.map((form) =>
      xmlElement(
        "form:form",
        { "form:name": form.name },
        form.controls.map((control) => formControlXml(control)),
      ),
    ),
  );
}

export function formControlXml(control: FormControlOptions): string {
  const common = {
    "form:id": control.id,
    "form:name": control.name,
    "form:disabled": control.disabled,
    "form:tab-index": control.tabIndex,
    "form:automatic-focus": control.automaticFocus,
  };
  if (control.control === "checkBox")
    return xmlElement("form:checkbox", {
      ...common,
      "form:current-state": control.checked ? "checked" : "unchecked",
    });
  if (control.control === "dropDownList")
    return xmlElement(
      "form:listbox",
      common,
      control.entries.map((label, optionIndex) =>
        xmlElement(
          "form:option",
          { "form:selected": control.selectedIndex === optionIndex || undefined },
          [escapeText(label)],
        ),
      ),
    );
  return xmlElement("form:text", {
    ...common,
    "office:value-type": control.valueType,
    "form:current-value": control.value,
    "form:max-length": control.maxLength,
  });
}

/** Rejects unrecognized slide content instead of emitting generic XML. */
export function assertElementNames(
  element: Element,
  allowedAttributes: string[],
  allowedChildren: string[] = [],
): void {
  const unknown = Object.keys(element.attributes ?? {}).filter(
    (name) => !allowedAttributes.includes(name),
  );
  if (unknown.length) throw unknownSlideChild(element, "/draw:page/office:forms");
  for (const child of element.elements ?? []) {
    if (child.type === "element" && !allowedChildren.includes(child.name ?? ""))
      throw unknownSlideChild(child, "/draw:page/office:forms");
  }
}

export function parseForms(page: Element): FormContainerOptions[] {
  const element = childNamed(page, "office:forms");
  if (!element) return [];
  assertElementNames(
    element,
    ["form:automatic-focus", "form:apply-design-mode"],
    ["form:form", "office:forms"],
  );
  return descendantElements(element)
    .filter((form) => form.name === "form:form")
    .map((form) => {
      assertElementNames(
        form,
        ["form:name"],
        ["form:text", "form:checkbox", "form:listbox", "form:form"],
      );
      return {
        ...(attributeString(form, "form:name") ? { name: attributeString(form, "form:name") } : {}),
        ...(attributeString(element, "form:automatic-focus") === "true"
          ? { automaticFocus: true }
          : {}),
        ...(attributeString(element, "form:apply-design-mode") === "true"
          ? { designMode: true }
          : {}),
        controls:
          form.elements
            ?.filter((child) => child.type === "element" && child.name !== "form:form")
            .map((control) => parseFormControl(control)) ?? [],
      };
    });
}

export function descendantElements(element: Element): Element[] {
  return (element.elements ?? []).flatMap((child) =>
    child.type === "element" ? [child, ...descendantElements(child)] : [],
  );
}

export function parseFormControl(control: Element): FormControlOptions {
  assertElementNames(
    control,
    [
      "form:id",
      "form:name",
      "form:disabled",
      "form:tab-index",
      "office:value-type",
      "form:automatic-focus",
      "form:current-value",
      "form:max-length",
      "form:current-state",
      "form:selected",
    ],
    ["form:option", "form:item"],
  );
  const common = {
    ...(attributeString(control, "form:id") ? { id: attributeString(control, "form:id") } : {}),
    ...(attributeString(control, "form:name")
      ? { name: attributeString(control, "form:name") }
      : {}),
    ...(attributeString(control, "form:disabled") === "true" ? { disabled: true } : {}),
    ...(attributeNumber(control, "form:tab-index") !== undefined
      ? { tabIndex: attributeNumber(control, "form:tab-index") }
      : {}),
    ...(attributeString(control, "form:automatic-focus") === "true"
      ? { automaticFocus: true }
      : {}),
  };
  if (control.name === "form:text")
    return {
      control: "text",
      ...common,
      ...(attributeString(control, "form:current-value")
        ? { value: attributeString(control, "form:current-value") }
        : {}),
      ...(attributeString(control, "office:value-type")
        ? { valueType: attributeString(control, "office:value-type") as "float" | "string" }
        : {}),
      ...(attributeNumber(control, "form:max-length") !== undefined
        ? { maxLength: attributeNumber(control, "form:max-length") }
        : {}),
    };
  if (control.name === "form:checkbox") {
    const state = attributeString(control, "form:current-state");
    if (state !== "checked" && state !== "unchecked")
      throw unknownSlideChild(control, "/office:forms");
    return { control: "checkBox", ...common, checked: state === "checked" };
  }
  if (control.name === "form:listbox") {
    const entries = childrenNamed(control, "form:option").map((option) => textOf(option));
    const selected = childrenNamed(control, "form:option").findIndex(
      (option) => attributeString(option, "form:selected") === "true",
    );
    return {
      control: "dropDownList",
      ...common,
      entries,
      ...(selected >= 0 ? { selectedIndex: selected } : {}),
    };
  }
  throw unknownSlideChild(control, "/office:forms");
}

export function unknownSlideChild(element: Element, parentPath: string): OdpParseError {
  const name = element.name ?? "(unknown)";
  const path = `${parentPath}/${name}`;
  const reason = "element has no canonical PresentationOptions mapping";
  return new OdpParseError(
    `content.xml: ${path}: ${name}: ${reason}`,
    "content.xml",
    path,
    name,
    reason,
  );
}
