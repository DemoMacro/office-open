/**
 * Descriptor registry for coverage tracking.
 *
 * Registers all descriptors so coverage tools can query which XML elements
 * are implemented. Replaces the old regex-based XSD scanning approach.
 *
 * @module
 */

import type { Descriptor } from "./types";

export class DescriptorRegistry {
  private static readonly _map = new Map<string, Descriptor<any>>();

  /** Register a descriptor with its XML tag. */
  public static register(tag: string, desc: Descriptor<any>): void {
    DescriptorRegistry._map.set(tag, desc);
  }

  /** Look up a descriptor by XML tag. */
  public static get(tag: string): Descriptor<any> | undefined {
    return DescriptorRegistry._map.get(tag);
  }

  /** Get all registered tags. */
  public static tags(): ReadonlySet<string> {
    return new Set(DescriptorRegistry._map.keys());
  }

  /** Check if a tag is registered. */
  public static has(tag: string): boolean {
    return DescriptorRegistry._map.has(tag);
  }

  /** Get the number of registered descriptors. */
  public static get size(): number {
    return DescriptorRegistry._map.size;
  }
}
