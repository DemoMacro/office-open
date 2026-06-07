/**
 * Diagram relationship IDs element.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd, CT_RelIds
 *
 * @module
 */
import { BuilderElement, type XmlComponent } from "../../xml-components";

// ---------------------------------------------------------------------------
// dgm:relIds — relationship IDs (CT_RelIds)
// ---------------------------------------------------------------------------

export interface DiagramRelIdsOptions {
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
export const createDiagramRelIds = (options: DiagramRelIdsOptions): XmlComponent =>
  new BuilderElement({
    name: "dgm:relIds",
    attributes: {
      "r:dm": { key: "r:dm", value: options.dm },
      "r:lo": { key: "r:lo", value: options.lo },
      "r:qs": { key: "r:qs", value: options.qs },
      "r:cs": { key: "r:cs", value: options.cs },
    },
  });
