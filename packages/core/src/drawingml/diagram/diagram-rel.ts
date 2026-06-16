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
  /** Relationship to data model part */
  dm: string;
  /** Relationship to layout definition part */
  lo: string;
  /** Relationship to quick style part */
  qs: string;
  /** Relationship to color transform part */
  cs: string;
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
  `<dgm:relIds r:dm="${options.dm}" r:lo="${options.lo}" r:qs="${options.qs}" r:cs="${options.cs}"/>`;
