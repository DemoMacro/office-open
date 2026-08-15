import { pickNonVisualDrawingProperties } from "../drawing";
import type {
  ConnectorLockingOptions,
  EndpointConnectionOptions,
  NonVisualDrawingPropertiesOptions,
} from "../drawing";

/**
 * Base connector options — the shared shape across pptx (p:cxnSp) and xlsx
 * (xdr:cxnSp) connectors.
 *
 * Carries the non-visual drawing properties (name/description/title/hidden that
 * mirror a:CT_NonVisualDrawingProps) plus the connector-specific glue: locking
 * (a:cxnSpLocks) and the optional endpoint connections (a:stCxn/a:endCxn). The
 * line geometry, fill, and outline stay package-side (pptx surfaces them as
 * top-level convenience fields, xlsx nests them in spPr), as does positioning
 * (pptx two endpoints x1/y1/x2/y2, xlsx cell anchor).
 *
 * docx has no standalone connector element, so it does not extend this base.
 */
export interface BaseConnectorOptions extends NonVisualDrawingPropertiesOptions {
  /** a:cxnSpLocks — connector locking (inside cNvCxnSpPr). */
  locking?: ConnectorLockingOptions;
  /** a:stCxn — start endpoint glued to a shape connection site. */
  startConnection?: EndpointConnectionOptions;
  /** a:endCxn — end endpoint glued to a shape connection site. */
  endConnection?: EndpointConnectionOptions;
}

/**
 * Pick the connector base fields (cNvPr + locking + endpoint connections)
 * actually set on `opts`, dropping undefined. Used when bridging a package's
 * connector options onto the shared base during cross-format conversion.
 */
export function pickConnectorBase(
  opts: BaseConnectorOptions | undefined,
): Partial<BaseConnectorOptions> {
  if (!opts) return {};
  return {
    ...pickNonVisualDrawingProperties(opts),
    ...(opts.locking !== undefined ? { locking: opts.locking } : {}),
    ...(opts.startConnection !== undefined ? { startConnection: opts.startConnection } : {}),
    ...(opts.endConnection !== undefined ? { endConnection: opts.endConnection } : {}),
  };
}
