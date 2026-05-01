/**
 * Table style override module for WordprocessingML documents.
 *
 * Table style overrides (CT_TblStylePr) allow defining formatting
 * for specific table regions (first row, last column, banding, etc.)
 * within a table style definition.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_TblStylePr
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Table style override type (ST_TblStyleOverrideType).
 *
 * Specifies which region of the table the override applies to.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TblStyleOverrideType">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="wholeTable"/>
 *     <xsd:enumeration value="firstRow"/>
 *     <xsd:enumeration value="lastRow"/>
 *     <xsd:enumeration value="firstCol"/>
 *     <xsd:enumeration value="lastCol"/>
 *     <xsd:enumeration value="band1Vert"/>
 *     <xsd:enumeration value="band2Vert"/>
 *     <xsd:enumeration value="band1Horz"/>
 *     <xsd:enumeration value="band2Horz"/>
 *     <xsd:enumeration value="neCell"/>
 *     <xsd:enumeration value="nwCell"/>
 *     <xsd:enumeration value="seCell"/>
 *     <xsd:enumeration value="swCell"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const TableStyleOverrideType = {
    /** Override applies to the entire table */
    WHOLE_TABLE: "wholeTable",
    /** Override applies to the first row */
    FIRST_ROW: "firstRow",
    /** Override applies to the last row */
    LAST_ROW: "lastRow",
    /** Override applies to the first column */
    FIRST_COL: "firstCol",
    /** Override applies to the last column */
    LAST_COL: "lastCol",
    /** Override applies to odd-numbered vertical bands (columns) */
    BAND1_VERT: "band1Vert",
    /** Override applies to even-numbered vertical bands (columns) */
    BAND2_VERT: "band2Vert",
    /** Override applies to odd-numbered horizontal bands (rows) */
    BAND1_HORZ: "band1Horz",
    /** Override applies to even-numbered horizontal bands (rows) */
    BAND2_HORZ: "band2Horz",
    /** Override applies to the northeast (top-right) corner cell */
    NE_CELL: "neCell",
    /** Override applies to the northwest (top-left) corner cell */
    NW_CELL: "nwCell",
    /** Override applies to the southeast (bottom-right) corner cell */
    SE_CELL: "seCell",
    /** Override applies to the southwest (bottom-left) corner cell */
    SW_CELL: "swCell",
} as const;

/**
 * Options for a table style override.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TblStylePr">
 *   <xsd:sequence>
 *     <xsd:element name="pPr" type="CT_PPrGeneral" minOccurs="0"/>
 *     <xsd:element name="rPr" type="CT_RPr" minOccurs="0"/>
 *     <xsd:element name="tblPr" type="CT_TblPrBase" minOccurs="0"/>
 *     <xsd:element name="trPr" type="CT_TrPr" minOccurs="0"/>
 *     <xsd:element name="tcPr" type="CT_TcPr" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="type" type="ST_TblStyleOverrideType" use="required"/>
 * </xsd:complexType>
 * ```
 */
export interface TableStyleOverrideOptions {
    /**
     * The table region this override applies to.
     */
    readonly type: (typeof TableStyleOverrideType)[keyof typeof TableStyleOverrideType];
    /**
     * Paragraph properties for this region.
     *
     * Accepts any XmlComponent (e.g., ParagraphProperties instance).
     */
    readonly paragraphProperties?: XmlComponent;
    /**
     * Run (character) properties for this region.
     *
     * Accepts any XmlComponent (e.g., RunProperties instance).
     */
    readonly runProperties?: XmlComponent;
    /**
     * Table properties for this region.
     *
     * Accepts any XmlComponent (e.g., TableProperties instance).
     */
    readonly tableProperties?: XmlComponent;
    /**
     * Table row properties for this region.
     *
     * Accepts any XmlComponent (e.g., TableRowProperties instance).
     */
    readonly rowProperties?: XmlComponent;
    /**
     * Table cell properties for this region.
     *
     * Accepts any XmlComponent (e.g., TableCellProperties instance).
     */
    readonly cellProperties?: XmlComponent;
}

/**
 * Creates a table style override element (w:tblStylePr).
 *
 * Table style overrides allow defining formatting for specific regions
 * of a table (first row, last column, banding, etc.) within a table
 * style definition.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TblStylePr">
 *   <xsd:sequence>
 *     <xsd:element name="pPr" type="CT_PPrGeneral" minOccurs="0"/>
 *     <xsd:element name="rPr" type="CT_RPr" minOccurs="0"/>
 *     <xsd:element name="tblPr" type="CT_TblPrBase" minOccurs="0"/>
 *     <xsd:element name="trPr" type="CT_TrPr" minOccurs="0"/>
 *     <xsd:element name="tcPr" type="CT_TcPr" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="type" type="ST_TblStyleOverrideType" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Bold text in the first row
 * createTableStyleOverride({
 *   type: "firstRow",
 *   runProperties: new RunProperties({ bold: true }),
 * });
 *
 * // Banded rows with background color
 * createTableStyleOverride({
 *   type: "band1Horz",
 *   cellProperties: new TableCellProperties({
 *     shading: { fill: "E7E6E6" },
 *   }),
 * });
 * ```
 */
export const createTableStyleOverride = (options: TableStyleOverrideOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.paragraphProperties) {
        children.push(options.paragraphProperties);
    }
    if (options.runProperties) {
        children.push(options.runProperties);
    }
    if (options.tableProperties) {
        children.push(options.tableProperties);
    }
    if (options.rowProperties) {
        children.push(options.rowProperties);
    }
    if (options.cellProperties) {
        children.push(options.cellProperties);
    }

    return new BuilderElement<{ readonly type: string }>({
        attributes: {
            type: { key: "w:type", value: options.type },
        },
        children,
        name: "w:tblStylePr",
    });
};
