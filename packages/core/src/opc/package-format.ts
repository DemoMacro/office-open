/** OOXML document family. Each family has one public root Options model. */
export type OoxmlPackageFamily = "wordprocessing" | "spreadsheet" | "presentation";

/** Package dialect emitted inside one OOXML family. */
export type OoxmlPackageVariant = "standard" | "macro" | "template" | "macroTemplate";

/** Concrete OOXML package format. */
export type OoxmlPackageFormat =
  | "docx"
  | "docm"
  | "dotx"
  | "dotm"
  | "xlsx"
  | "xlsm"
  | "xltx"
  | "xltm"
  | "pptx"
  | "pptm"
  | "potx"
  | "potm";

export interface OoxmlPackageFormatInfo {
  family: OoxmlPackageFamily;
  variant: OoxmlPackageVariant;
  mimeType: string;
  mainPartPath: string;
  mainContentType: string;
}

const WORD_MAIN_PART = "word/document.xml";
const SPREADSHEET_MAIN_PART = "xl/workbook.xml";
const PRESENTATION_MAIN_PART = "ppt/presentation.xml";

export const OOXML_PACKAGE_FORMATS: Record<OoxmlPackageFormat, OoxmlPackageFormatInfo> = {
  docx: {
    family: "wordprocessing",
    variant: "standard",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    mainPartPath: WORD_MAIN_PART,
    mainContentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  },
  docm: {
    family: "wordprocessing",
    variant: "macro",
    mimeType: "application/vnd.ms-word.document.macroEnabled.12",
    mainPartPath: WORD_MAIN_PART,
    mainContentType: "application/vnd.ms-word.document.macroEnabled.main+xml",
  },
  dotx: {
    family: "wordprocessing",
    variant: "template",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    mainPartPath: WORD_MAIN_PART,
    mainContentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
  },
  dotm: {
    family: "wordprocessing",
    variant: "macroTemplate",
    mimeType: "application/vnd.ms-word.template.macroEnabled.12",
    mainPartPath: WORD_MAIN_PART,
    mainContentType: "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
  },
  xlsx: {
    family: "spreadsheet",
    variant: "standard",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    mainPartPath: SPREADSHEET_MAIN_PART,
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  },
  xlsm: {
    family: "spreadsheet",
    variant: "macro",
    mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12",
    mainPartPath: SPREADSHEET_MAIN_PART,
    mainContentType: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  },
  xltx: {
    family: "spreadsheet",
    variant: "template",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    mainPartPath: SPREADSHEET_MAIN_PART,
    mainContentType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
  },
  xltm: {
    family: "spreadsheet",
    variant: "macroTemplate",
    mimeType: "application/vnd.ms-excel.template.macroEnabled.12",
    mainPartPath: SPREADSHEET_MAIN_PART,
    mainContentType: "application/vnd.ms-excel.template.macroEnabled.main+xml",
  },
  pptx: {
    family: "presentation",
    variant: "standard",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    mainPartPath: PRESENTATION_MAIN_PART,
    mainContentType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  },
  pptm: {
    family: "presentation",
    variant: "macro",
    mimeType: "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
    mainPartPath: PRESENTATION_MAIN_PART,
    mainContentType: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
  },
  potx: {
    family: "presentation",
    variant: "template",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.template",
    mainPartPath: PRESENTATION_MAIN_PART,
    mainContentType:
      "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
  },
  potm: {
    family: "presentation",
    variant: "macroTemplate",
    mimeType: "application/vnd.ms-powerpoint.template.macroEnabled.12",
    mainPartPath: PRESENTATION_MAIN_PART,
    mainContentType: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
  },
};

export function ooxmlPackageFormatInfo(
  family: OoxmlPackageFamily,
  variant: OoxmlPackageVariant = "standard",
): OoxmlPackageFormatInfo {
  const match = Object.values(OOXML_PACKAGE_FORMATS).find(
    (format) => format.family === family && format.variant === variant,
  );
  if (!match) throw new Error(`Unknown OOXML package variant: ${family}/${variant}`);
  return match;
}

/**
 * Main-part Override content type for the output package. The preserved
 * source declaration wins so a dotx round-trip keeps template.main; callers
 * re-typing a package (aggregate dispatch to an explicit format) rewrite the
 * declaration before compiling. Fresh compiles without a source table fall
 * back to the requested variant's main content type.
 */
export function resolveMainPartContentType(
  family: OoxmlPackageFamily,
  source: { overrides: ReadonlyArray<{ partName: string; contentType: string }> } | undefined,
  mainPartPath: string,
  fallback: string,
): string {
  const partName = `/${mainPartPath}`.toLowerCase();
  const declared = source?.overrides.find(
    (override) => override.partName.toLowerCase() === partName,
  )?.contentType;
  if (!declared) return fallback;
  const isFamilyMain = Object.values(OOXML_PACKAGE_FORMATS).some(
    (format) => format.family === family && format.mainContentType === declared,
  );
  return isFamilyMain ? declared : fallback;
}
