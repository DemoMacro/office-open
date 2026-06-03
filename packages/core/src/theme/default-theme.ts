/**
 * DefaultTheme — cached theme XML component.
 *
 * - No-arg constructor returns module-level pre-computed constant (zero allocation).
 * - With options, builds XML string and caches by key.
 *
 * @module
 */
import { BaseXmlComponent, type Context } from "../xml-components";
import { buildThemeXml } from "./build-theme-xml";
import type { ThemeOptions } from "./theme-options";

// Pre-computed default theme XML — zero allocation for the common case.
const DEFAULT_XML = buildThemeXml();

// Cache for customized themes, keyed by serialized options.
const customCache = new Map<string, string>();

function themeKey(o: ThemeOptions): string {
  const c = o.colors;
  const f = o.fonts;
  return `n${o.name ?? ""}c${c?.dark1 ?? ""}${c?.light1 ?? ""}${c?.dark2 ?? ""}${c?.light2 ?? ""}${c?.accent1 ?? ""}${c?.accent2 ?? ""}${c?.accent3 ?? ""}${c?.accent4 ?? ""}${c?.accent5 ?? ""}${c?.accent6 ?? ""}${c?.hyperlink ?? ""}${c?.followedHyperlink ?? ""}f${f?.majorFont ?? ""}${f?.minorFont ?? ""}${f?.majorFontAsian ?? ""}${f?.minorFontAsian ?? ""}`;
}

export class DefaultTheme extends BaseXmlComponent {
  private readonly xml: string;

  public constructor(options?: ThemeOptions) {
    super("a:theme");
    if (!options) {
      this.xml = DEFAULT_XML;
      return;
    }
    const key = themeKey(options);
    const cached = customCache.get(key);
    if (cached) {
      this.xml = cached;
    } else {
      const built = buildThemeXml(options);
      customCache.set(key, built);
      this.xml = built;
    }
  }

  public override toXml(_context: Context): string {
    return this.xml;
  }
}
