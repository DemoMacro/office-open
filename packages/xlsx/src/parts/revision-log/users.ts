/**
 * Users — xl/users.xml (CT_Users, sml.xsd:2100) descriptor.
 *
 * Shared-workbook user list.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, children, escapeXml } from "@office-open/xml";

import { S_NS } from "./stringify";
import type { UsersOptions } from "./types";

export const usersDesc: CustomDescriptor<UsersOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (!opts.users || opts.users.length === 0) return undefined;
    const users = opts.users
      .map(
        (u) =>
          `<userInfo guid="${escapeXml(u.guid)}" name="${escapeXml(u.name)}" id="${u.id}"` +
          ` dateTime="${escapeXml(u.dateTime)}"/>`,
      )
      .join("");
    return `<users xmlns="${S_NS}" count="${opts.users.length}">${users}</users>`;
  },

  parse(el, _ctx) {
    const users = children(el, "userInfo").map((u) => ({
      guid: attr(u, "guid") ?? "",
      name: attr(u, "name") ?? "",
      id: Number(attr(u, "id") ?? "0"),
      dateTime: attr(u, "dateTime") ?? "",
    }));
    const result: Partial<UsersOptions> = {};
    if (users.length > 0) result.users = users;
    return result as UsersOptions;
  },
};
