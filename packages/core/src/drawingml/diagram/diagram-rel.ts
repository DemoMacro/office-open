/**
 * Diagram relationship IDs element.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd, CT_RelIds
 *
 * @module
 */

// ---------------------------------------------------------------------------
// dgm:relIds — relationship IDs (CT_RelIds)
// ---------------------------------------------------------------------------

export interface DiagramRelationshipIdsOptions {
  /** Relationship to data model part (@r:dm) */
  dataModel: string;
  /** Relationship to layout definition part (@r:lo) */
  layout: string;
  /** Relationship to quick style part (@r:qs) */
  quickStyle: string;
  /** Relationship to color transform part (@r:cs) */
  colorStyle: string;
}

/**
 * Creates a dgm:relIds element (CT_RelIds).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_RelIds">
 *   <xsd:attribute ref="r:dm" use="required"/>
 *   <xsd:attribute ref="r:lo" use="required"/>
 *   <xsd:attribute ref="r:qs" use="required"/>
 *   <xsd:attribute ref="r:cs" use="required"/>
 * </xsd:complexType>
 * ```
 */
export const createDiagramRelationshipIds = (options: DiagramRelationshipIdsOptions): string =>
  `<dgm:relIds r:dm="${options.dataModel}" r:lo="${options.layout}" r:qs="${options.quickStyle}" r:cs="${options.colorStyle}"/>`;
