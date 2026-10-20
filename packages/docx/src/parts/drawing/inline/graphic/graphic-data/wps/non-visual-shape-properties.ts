import type {
  ConnectorLockingOptions,
  NonVisualDrawingPropertiesOptions,
  EndpointConnectionOptions,
  ShapeLockingOptions,
} from "@office-open/core/drawing";

import type { HyperlinkOptions } from "../../../../doc-properties/doc-properties";

export interface NonVisualShapePropertiesOptions extends NonVisualDrawingPropertiesOptions {
  /** Drawing id (wps:cNvPr `@id`). */
  id?: number;
  /** Click or hover hyperlink on the shape itself (wps:cNvPr/a:hlinkClick). */
  hyperlink?: HyperlinkOptions;
  /** Shape type marker (wps:cNvSpPr `@txBox`) — "1" marks a text box. */
  textBox?: string;
  /**
   * Connector shape marker — when true the shape emits wps:cNvCnPr (connector
   * non-visual properties) instead of wps:cNvSpPr. Per the CT_WordprocessingShape
   * choice these are mutually exclusive.
   */
  connector?: boolean;
  /** Shape locks (wps:cNvSpPr/a:spLocks). An empty object = bare element. */
  locking?: ShapeLockingOptions;
  /** Connector locks (wps:cNvCnPr/a:cxnSpLocks), when connector is set. */
  connectorLocking?: ConnectorLockingOptions;
  /** Connector start endpoint (wps:cNvCnPr/a:stCxn). */
  startConnection?: EndpointConnectionOptions;
  /** Connector end endpoint (wps:cNvCnPr/a:endCxn). */
  endConnection?: EndpointConnectionOptions;
}
