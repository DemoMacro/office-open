/**
 * Test utility module.
 *
 * Provides helper functions for unit tests.
 *
 * @module
 * @internal
 */

/**
 * Utility class for test helpers.
 *
 * @internal
 */
export class Utility {
    public static jsonify(obj: object): any {
        const stringifiedJson = JSON.stringify(obj);
        return JSON.parse(stringifiedJson);
    }
}
