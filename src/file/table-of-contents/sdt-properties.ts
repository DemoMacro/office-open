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
import {
    BuilderElement,
    OnOffElement,
    StringValueElement,
    XmlComponent,
} from "@file/xml-components";
import type { XmlComponent as XmlComponentType } from "@file/xml-components";

// ─── Lock ───────────────────────────────────────────────────────────────────

/**
 * SDT lock type (ST_Lock).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_Lock">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="sdtLocked"/>
 *     <xsd:enumeration value="contentLocked"/>
 *     <xsd:enumeration value="unlocked"/>
 *     <xsd:enumeration value="sdtContentLocked"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
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
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtListItem">
 *   <xsd:attribute name="displayText" type="s:ST_String"/>
 *   <xsd:attribute name="value" type="s:ST_String"/>
 * </xsd:complexType>
 * ```
 */
export interface SdtListItem {
    /** Display text shown in the UI */
    readonly displayText?: string;
    /** Underlying value */
    readonly value?: string;
}

/**
 * Options for comboBox SDT type (CT_SdtComboBox).
 *
 * A comboBox allows users to either select from a list or type custom text.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtComboBox">
 *   <xsd:sequence>
 *     <xsd:element name="listItem" type="CT_SdtListItem" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="lastValue" type="s:ST_String" use="optional" default=""/>
 * </xsd:complexType>
 * ```
 */
export interface SdtComboBoxOptions {
    /** List items */
    readonly items?: readonly SdtListItem[];
    /** Last selected value */
    readonly lastValue?: string;
}

/**
 * Options for dropDownList SDT type (CT_SdtDropDownList).
 *
 * A dropDownList restricts users to selecting from a predefined list.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtDropDownList">
 *   <xsd:sequence>
 *     <xsd:element name="listItem" type="CT_SdtListItem" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="lastValue" type="s:ST_String" use="optional" default=""/>
 * </xsd:complexType>
 * ```
 */
export interface SdtDropDownListOptions {
    /** List items */
    readonly items?: readonly SdtListItem[];
    /** Last selected value */
    readonly lastValue?: string;
}

/**
 * Date mapping type (ST_SdtDateMappingType).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_SdtDateMappingType">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="text"/>
 *     <xsd:enumeration value="date"/>
 *     <xsd:enumeration value="dateTime"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const SdtDateMappingType = {
    TEXT: "text",
    DATE: "date",
    DATE_TIME: "dateTime",
} as const;

/**
 * Options for date SDT type (CT_SdtDate).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtDate">
 *   <xsd:sequence>
 *     <xsd:element name="dateFormat" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="lid" type="CT_Lang" minOccurs="0"/>
 *     <xsd:element name="storeMappedDataAs" type="CT_SdtDateMappingType" minOccurs="0"/>
 *     <xsd:element name="calendar" type="CT_CalendarType" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="fullDate" type="ST_DateTime" use="optional"/>
 * </xsd:complexType>
 * ```
 */
export interface SdtDateOptions {
    /** Date format string (e.g., "yyyy-MM-dd") */
    readonly dateFormat?: string;
    /** Language ID (e.g., "en-US") */
    readonly languageId?: string;
    /** How the date value is stored */
    readonly storeMappedDataAs?: (typeof SdtDateMappingType)[keyof typeof SdtDateMappingType];
    /** Calendar type */
    readonly calendar?: string;
    /** Full date value (ISO 8601) */
    readonly fullDate?: string;
}

/**
 * Options for plain text SDT type (CT_SdtText).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtText">
 *   <xsd:attribute name="multiLine" type="s:ST_OnOff"/>
 * </xsd:complexType>
 * ```
 */
export interface SdtTextOptions {
    /** Whether the text control supports multiple lines */
    readonly multiLine?: boolean;
}

/**
 * Data binding options (CT_DataBinding).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_DataBinding">
 *   <xsd:attribute name="prefixMappings" type="s:ST_String"/>
 *   <xsd:attribute name="xpath" type="s:ST_String" use="required"/>
 *   <xsd:attribute name="storeItemID" type="s:ST_String" use="required"/>
 * </xsd:complexType>
 * ```
 */
export interface SdtDataBindingOptions {
    /** XML namespace prefix mappings */
    readonly prefixMappings?: string;
    /** XPath expression (required) */
    readonly xpath: string;
    /** Custom XML store item ID (required) */
    readonly storeItemID: string;
}

/**
 * Options for CT_SdtPr — structured document tag properties.
 *
 * Exactly zero or one type discriminator from the `xsd:choice` group may be specified:
 * `comboBox`, `dropDownList`, `date`, `text`, `equation`, `richText`, `picture`,
 * `citation`, `group`, `bibliography`, `docPartObj`, `docPartList`.
 */
export interface SdtPropertiesOptions {
    /** Display name */
    readonly alias?: string;
    /** Application-specific tag */
    readonly tag?: string;
    /** Unique ID */
    readonly id?: number;
    /** Lock behavior */
    readonly lock?: (typeof SdtLock)[keyof typeof SdtLock];
    /** Whether the control is temporary */
    readonly temporary?: boolean;
    /** Whether the placeholder text is currently shown */
    readonly showingPlaceholder?: boolean;
    /** Placeholder content (block-level content: paragraphs, tables, etc.) */
    readonly placeholder?: XmlComponentType[];
    /** Data binding to custom XML */
    readonly dataBinding?: SdtDataBindingOptions;
    /** Numeric label */
    readonly label?: number;
    /** Tab order index */
    readonly tabIndex?: number;
    /** SDT content run properties (rPr) */
    readonly runProperties?: XmlComponentType;

    // ─── Type discriminators (xsd:choice, at most one) ───

    /** Equation SDT */
    readonly equation?: boolean;
    /** ComboBox SDT (allows free-text entry) */
    readonly comboBox?: SdtComboBoxOptions;
    /** Date SDT */
    readonly date?: SdtDateOptions;
    /** Document part object SDT */
    readonly docPartObj?: {
        readonly gallery?: string;
        readonly category?: string;
        readonly unique?: boolean;
    };
    /** Document part list SDT */
    readonly docPartList?: {
        readonly gallery?: string;
        readonly category?: string;
        readonly unique?: boolean;
    };
    /** DropDownList SDT (selection only) */
    readonly dropDownList?: SdtDropDownListOptions;
    /** Picture SDT */
    readonly picture?: boolean;
    /** Rich text SDT */
    readonly richText?: boolean;
    /** Plain text SDT */
    readonly text?: SdtTextOptions;
    /** Citation SDT */
    readonly citation?: boolean;
    /** Group SDT */
    readonly group?: boolean;
    /** Bibliography SDT */
    readonly bibliography?: boolean;
}

// ─── Helpers ────────────────────────────────────────────────────────────────

/** Creates a list item element (w:listItem).
 *
 * Note: Word requires `w:value` on every listItem inside w:dropDownList,
 * even though the XSD marks it as optional. When value is not provided,
 * it defaults to displayText.
 */
const createListItem = (item: SdtListItem, forceValue?: boolean): XmlComponentType => {
    const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
    if (item.displayText !== undefined) {
        attrs.displayText = { key: "w:displayText", value: item.displayText };
    }
    const value = item.value ?? (forceValue ? item.displayText : undefined);
    if (value !== undefined) {
        attrs.value = { key: "w:value", value };
    }
    return new BuilderElement({ name: "w:listItem", attributes: attrs });
};

/** Creates a comboBox or dropDownList element. */
const createListType = (
    name: string,
    options: { readonly items?: readonly SdtListItem[]; readonly lastValue?: string },
): XmlComponentType => {
    const children: XmlComponentType[] = [];
    if (options.items) {
        for (const item of options.items) {
            children.push(createListItem(item, name === "w:dropDownList"));
        }
    }
    const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
    if (options.lastValue !== undefined) {
        attrs.lastValue = { key: "w:lastValue", value: options.lastValue };
    }
    return new BuilderElement({
        name,
        attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
        children: children.length > 0 ? children : undefined,
    });
};

/** Creates a date SDT element (w:date). */
const createDate = (options: SdtDateOptions): XmlComponentType => {
    const children: XmlComponentType[] = [];
    if (options.dateFormat !== undefined) {
        children.push(new StringValueElement("w:dateFormat", options.dateFormat));
    }
    if (options.languageId !== undefined) {
        children.push(new StringValueElement("w:lid", options.languageId));
    }
    if (options.storeMappedDataAs !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                name: "w:storeMappedDataAs",
                attributes: { val: { key: "w:val", value: options.storeMappedDataAs } },
            }),
        );
    }
    if (options.calendar !== undefined) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                name: "w:calendar",
                attributes: { val: { key: "w:val", value: options.calendar } },
            }),
        );
    }
    const attrs: Record<string, { readonly key: string; readonly value: string }> = {};
    if (options.fullDate !== undefined) {
        attrs.fullDate = { key: "w:fullDate", value: options.fullDate };
    }
    return new BuilderElement({
        name: "w:date",
        attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
        children: children.length > 0 ? children : undefined,
    });
};

/** Creates a dataBinding element (w:dataBinding). */
const createDataBinding = (options: SdtDataBindingOptions): XmlComponentType => {
    const attrs: Record<string, { readonly key: string; readonly value: string }> = {
        xpath: { key: "w:xpath", value: options.xpath },
        storeItemID: { key: "w:storeItemID", value: options.storeItemID },
    };
    if (options.prefixMappings !== undefined) {
        attrs.prefixMappings = { key: "w:prefixMappings", value: options.prefixMappings };
    }
    return new BuilderElement({ name: "w:dataBinding", attributes: attrs });
};

/** Creates a docPart element (w:docPartObj or w:docPartList). */
const createDocPart = (
    name: string,
    options: { readonly gallery?: string; readonly category?: string; readonly unique?: boolean },
): XmlComponentType => {
    const children: XmlComponentType[] = [];
    if (options.gallery !== undefined) {
        children.push(new StringValueElement("w:docPartGallery", options.gallery));
    }
    if (options.category !== undefined) {
        children.push(new StringValueElement("w:docPartCategory", options.category));
    }
    if (options.unique !== undefined) {
        children.push(new OnOffElement("w:docPartUnique", options.unique));
    }
    return new BuilderElement({ name, children: children.length > 0 ? children : undefined });
};

// ─── Main Class ─────────────────────────────────────────────────────────────

/**
 * Represents the properties of a Structured Document Tag (CT_SdtPr).
 *
 * Supports the full CT_SdtPr schema including all SDT type discriminators
 * (comboBox, dropDownList, date, text, etc.) defined in the `xsd:choice` group.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtPr">
 *   <xsd:sequence>
 *     <xsd:element name="rPr" type="CT_RPr" minOccurs="0"/>
 *     <xsd:element name="alias" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="tag" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="id" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="lock" type="CT_Lock" minOccurs="0"/>
 *     <xsd:element name="placeholder" type="CT_Placeholder" minOccurs="0"/>
 *     <xsd:element name="temporary" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="showingPlcHdr" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="dataBinding" type="CT_DataBinding" minOccurs="0"/>
 *     <xsd:element name="label" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="tabIndex" type="CT_UnsignedDecimalNumber" minOccurs="0"/>
 *     <xsd:choice minOccurs="0" maxOccurs="1">
 *       <xsd:element name="equation" type="CT_Empty"/>
 *       <xsd:element name="comboBox" type="CT_SdtComboBox"/>
 *       <xsd:element name="date" type="CT_SdtDate"/>
 *       <xsd:element name="docPartObj" type="CT_SdtDocPart"/>
 *       <xsd:element name="docPartList" type="CT_SdtDocPart"/>
 *       <xsd:element name="dropDownList" type="CT_SdtDropDownList"/>
 *       <xsd:element name="picture" type="CT_Empty"/>
 *       <xsd:element name="richText" type="CT_Empty"/>
 *       <xsd:element name="text" type="CT_SdtText"/>
 *       <xsd:element name="citation" type="CT_Empty"/>
 *       <xsd:element name="group" type="CT_Empty"/>
 *       <xsd:element name="bibliography" type="CT_Empty"/>
 *     </xsd:choice>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Simple alias (backward compatible)
 * new StructuredDocumentTagProperties("Table of Contents");
 *
 * // Full options with comboBox type
 * new StructuredDocumentTagProperties({
 *   alias: "Pick a color",
 *   comboBox: { items: [{ displayText: "Red", value: "r" }, { displayText: "Blue", value: "b" }] },
 * });
 * ```
 */
export class StructuredDocumentTagProperties extends XmlComponent {
    public constructor(aliasOrOptions?: string | SdtPropertiesOptions) {
        super("w:sdtPr");

        // Backward compatibility: accept a plain string alias
        const options: SdtPropertiesOptions | undefined =
            typeof aliasOrOptions === "string" ? { alias: aliasOrOptions } : aliasOrOptions;

        if (!options) {
            return;
        }

        // CT_SdtPr children in XSD sequence order
        if (options.runProperties) {
            this.root.push(options.runProperties);
        }
        if (options.alias !== undefined) {
            this.root.push(new StringValueElement("w:alias", options.alias));
        }
        if (options.tag !== undefined) {
            this.root.push(new StringValueElement("w:tag", options.tag));
        }
        if (options.id !== undefined) {
            this.root.push(
                new BuilderElement<{ readonly val: number }>({
                    name: "w:id",
                    attributes: { val: { key: "w:val", value: options.id } },
                }),
            );
        }
        if (options.lock !== undefined) {
            this.root.push(
                new BuilderElement<{ readonly val: string }>({
                    name: "w:lock",
                    attributes: { val: { key: "w:val", value: options.lock } },
                }),
            );
        }
        if (options.placeholder !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "w:placeholder",
                    children: options.placeholder.length > 0 ? options.placeholder : undefined,
                }),
            );
        }
        if (options.temporary !== undefined) {
            this.root.push(new OnOffElement("w:temporary", options.temporary));
        }
        // When placeholder has content, showingPlcHdr must be true for Word to render correctly
        const effectiveShowingPlcHdr =
            options.showingPlaceholder ??
            (options.placeholder !== undefined && options.placeholder.length > 0);
        if (options.showingPlaceholder !== undefined || effectiveShowingPlcHdr) {
            this.root.push(new OnOffElement("w:showingPlcHdr", effectiveShowingPlcHdr));
        }
        if (options.dataBinding) {
            this.root.push(createDataBinding(options.dataBinding));
        }
        if (options.label !== undefined) {
            this.root.push(
                new BuilderElement<{ readonly val: number }>({
                    name: "w:label",
                    attributes: { val: { key: "w:val", value: options.label } },
                }),
            );
        }
        if (options.tabIndex !== undefined) {
            this.root.push(
                new BuilderElement<{ readonly val: number }>({
                    name: "w:tabIndex",
                    attributes: { val: { key: "w:val", value: options.tabIndex } },
                }),
            );
        }

        // Type discriminator (xsd:choice — at most one)
        if (options.equation) {
            this.root.push(new BuilderElement({ name: "w:equation" }));
        } else if (options.comboBox) {
            this.root.push(createListType("w:comboBox", options.comboBox));
        } else if (options.date) {
            this.root.push(createDate(options.date));
        } else if (options.docPartObj) {
            this.root.push(createDocPart("w:docPartObj", options.docPartObj));
        } else if (options.docPartList) {
            this.root.push(createDocPart("w:docPartList", options.docPartList));
        } else if (options.dropDownList) {
            this.root.push(createListType("w:dropDownList", options.dropDownList));
        } else if (options.picture) {
            this.root.push(new BuilderElement({ name: "w:picture" }));
        } else if (options.richText) {
            this.root.push(new BuilderElement({ name: "w:richText" }));
        } else if (options.text !== undefined) {
            const textAttrs: Record<
                string,
                { readonly key: string; readonly value: string | boolean }
            > = {};
            if (options.text.multiLine !== undefined) {
                textAttrs.multiLine = { key: "w:multiLine", value: options.text.multiLine };
            } else {
                // Word requires at least one attribute on w:text when showingPlcHdr is true;
                // explicitly emit the default to avoid generating an empty <w:text/> element
                textAttrs.multiLine = { key: "w:multiLine", value: false };
            }
            this.root.push(
                new BuilderElement({
                    name: "w:text",
                    attributes: textAttrs,
                }),
            );
        } else if (options.citation) {
            this.root.push(new BuilderElement({ name: "w:citation" }));
        } else if (options.group) {
            this.root.push(new BuilderElement({ name: "w:group" }));
        } else if (options.bibliography) {
            this.root.push(new BuilderElement({ name: "w:bibliography" }));
        }
    }
}
