const RELATIONSHIP_ATTRIBUTE = /\br:(?:id|embed|link)="([^"]+)"/g;

export function collectRawXmlRelationshipIds(
  value: unknown,
  ids: Set<string> = new Set(),
): Set<string> {
  const pending: unknown[] = [value];
  while (pending.length > 0) {
    const current = pending.pop();
    if (Array.isArray(current)) {
      pending.push(...current);
      continue;
    }
    if (!current || typeof current !== "object") continue;
    for (const [key, child] of Object.entries(current)) {
      if (key === "rawXml" && typeof child === "string") {
        for (const match of child.matchAll(RELATIONSHIP_ATTRIBUTE)) {
          const id = match[1];
          if (id) ids.add(id);
        }
        continue;
      }
      pending.push(child);
    }
  }
  return ids;
}
