/**
 * Structured Document Tag Properties module.
 *
 * Represents CT_SdtPr — the properties of a structured document tag (content control).
 * Supports all SDT types defined in the OOXML specification: comboBox, dropDownList,
 * date, text, equation,richText, picture, citation, group, bibliography, docPartObj,
 * docPartList.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_SdtPr
 *
 * @module
 */

// ─── Lock ───────────────────────────────────────────────────────────────────

/**
 * SDT lock type (ST_Lock).
 */
export const SdtLock = {
  /** Lock the SDT itself (cannot delete the control) */
  SDT_LOCKED: "sdtLocked",
  /** Lock the content (cannot edit content within the control) */
  CONTENT_LOCKED: "contentLocked",
  /** No locking */
  UNLOCKED: "unlocked",
  /** Lock both SDT and content */
  SDT_CONTENT_LOCKED: "sdtContentLocked",
} as const;

// ─── Supporting Types ───────────────────────────────────────────────────────

/**
 * A list item for comboBox or dropDownList SDTs (CT_SdtListItem).
 */
export interface SdtListItem {
  /** Display text shown in the UI */
  displayText?: string;
  /** Underlying value */
  value?: string;
}

/**
 * Options for comboBox SDT type (CT_SdtComboBox).
 */
export interface SdtComboBoxOptions {
  /** List items */
  items?: SdtListItem[];
  /** Last selected value */
  lastValue?: string;
}

/**
 * Options for dropDownList SDT type (CT_SdtDropDownList).
 */
export interface SdtDropDownListOptions {
  /** List items */
  items?: SdtListItem[];
  /** Last selected value */
  lastValue?: string;
}

/**
 * Date mapping type (ST_SdtDateMappingType).
 */
export const SdtDateMappingType = {
  TEXT: "text",
  DATE: "date",
  DATE_TIME: "dateTime",
} as const;

/**
 * Options for date SDT type (CT_SdtDate).
 */
export interface SdtDateOptions {
  /** Date format string (e.g., "yyyy-MM-dd") */
  dateFormat?: string;
  /** Language ID (e.g., "en-US") */
  languageId?: string;
  /** How the date value is stored */
  storeMappedDataAs?: (typeof SdtDateMappingType)[keyof typeof SdtDateMappingType];
  /** Calendar type */
  calendar?: string;
  /** Full date value (ISO 8601) */
  fullDate?: string;
}

/**
 * Options for plain text SDT type (CT_SdtText).
 */
export interface SdtTextOptions {
  /** Whether the text control supports multiple lines */
  multiLine?: boolean;
}

/**
 * Data binding options (CT_DataBinding).
 */
export interface SdtDataBindingOptions {
  /** XML namespace prefix mappings */
  prefixMappings?: string;
  /** XPath expression (required) */
  xpath: string;
  /** Custom XML store item ID (required) */
  storeItemID: string;
}

/**
 * Options for CT_SdtPr — structured document tag properties.
 */
export interface SdtPropertiesOptions {
  /** Display name */
  alias?: string;
  /** Application-specific tag */
  tag?: string;
  /** Unique ID */
  id?: number;
  /** Lock behavior */
  lock?: (typeof SdtLock)[keyof typeof SdtLock];
  /** Whether the control is temporary */
  temporary?: boolean;
  /** Whether the placeholder text is currently shown */
  showingPlaceholder?: boolean;
  /** Data binding to custom XML */
  dataBinding?: SdtDataBindingOptions;
  /** Numeric label */
  label?: number;
  /** Tab order index */
  tabIndex?: number;

  // ─── Type discriminators (xsd:choice, at most one) ───

  /** Equation SDT */
  equation?: boolean;
  /** ComboBox SDT (allows free-text entry) */
  comboBox?: SdtComboBoxOptions;
  /** Date SDT */
  date?: SdtDateOptions;
  /** Document part object SDT */
  docPartObj?: {
    gallery?: string;
    category?: string;
    unique?: boolean;
  };
  /** Document part list SDT */
  docPartList?: {
    gallery?: string;
    category?: string;
    unique?: boolean;
  };
  /** DropDownList SDT (selection only) */
  dropDownList?: SdtDropDownListOptions;
  /** Picture SDT */
  picture?: boolean;
  /** Rich text SDT */
  richText?: boolean;
  /** Plain text SDT */
  text?: SdtTextOptions;
  /** Citation SDT */
  citation?: boolean;
  /** Group SDT */
  group?: boolean;
  /** Bibliography SDT */
  bibliography?: boolean;
}
