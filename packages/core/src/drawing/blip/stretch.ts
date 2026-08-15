/**
 * Stretch fill stringifier for blip fills.
 *
 * @module
 */

/**
 * Build a:stretch XML string with a:fillRect child.
 */
export function stringifyStretch(): string {
  return "<a:stretch><a:fillRect/></a:stretch>";
}
