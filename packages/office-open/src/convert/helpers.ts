import type { BaseConnectorOptions, BaseGroupOptions } from "@office-open/core";
import { pickNonVisualDrawingProperties } from "@office-open/core";

export function pickConnectorBase(
  options: BaseConnectorOptions | undefined,
): Partial<BaseConnectorOptions> {
  if (!options) return {};
  return {
    ...pickNonVisualDrawingProperties(options),
    ...(options.locking !== undefined ? { locking: options.locking } : {}),
    ...(options.startConnection !== undefined ? { startConnection: options.startConnection } : {}),
    ...(options.endConnection !== undefined ? { endConnection: options.endConnection } : {}),
  };
}

export function pickGroupBase(options: BaseGroupOptions | undefined): Partial<BaseGroupOptions> {
  if (!options) return {};
  return {
    ...pickNonVisualDrawingProperties(options),
    ...(options.locking !== undefined ? { locking: options.locking } : {}),
  };
}
