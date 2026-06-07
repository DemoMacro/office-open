/**
 * Slide synchronization property types.
 *
 * @module
 */

export interface SlideSyncOptions {
  readonly serverSldId: string;
  readonly serverSldModifiedTime: string;
  readonly clientInsertedTime: string;
}
