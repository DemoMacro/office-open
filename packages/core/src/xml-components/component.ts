/**
 * Core XML Component classes for building OOXML element trees.
 *
 * @module
 */
import { escapeXml, xml } from "@office-open/xml";

import { BaseXmlComponent } from "./base";
import type { Context, IXmlableObject } from "./base";

/**
 * Empty object singleton used for empty XML elements.
 *
 * @internal
 */
export const EMPTY_OBJECT = Object.seal({});

/**
 * Base class for all XML components in OOXML documents.
 *
 * `toXml()` traverses the `root` array and calls `toXml()` on each
 * `BaseXmlComponent` child — this avoids building the intermediate
 * `IXmlableObject` tree. Inline `IXmlableObject` children fall back to
 * `xml()` for serialization.
 */
export abstract class XmlComponent extends BaseXmlComponent {
  /**
   * Array of child components, text nodes, and attributes.
   */
  public root: (BaseXmlComponent | IXmlableObject | string)[];

  public constructor(rootKey: string) {
    super(rootKey);
    this.root = [];
  }

  /**
   * Prepares this component and its children for XML serialization.
   */
  public prepForXml(context: Context): IXmlableObject | undefined {
    context.stack.push(this);

    const children: (BaseXmlComponent | IXmlableObject | string | undefined)[] = [];
    for (const comp of this.root) {
      if (comp instanceof BaseXmlComponent) {
        const prepared = comp.prepForXml(context);
        if (prepared !== undefined) {
          children.push(prepared);
        }
      } else {
        children.push(comp);
      }
    }

    context.stack.pop();

    return {
      [this.rootKey]: children.length
        ? children.length === 1 &&
          children[0] &&
          typeof children[0] === "object" &&
          "_attr" in children[0]
          ? children[0]
          : children
        : EMPTY_OBJECT,
    };
  }

  /**
   * Direct XML serialization — traverses `root` and calls `toXml()` on
   * each child, concatenating the results into a single string.
   */
  public toXml(context: Context): string {
    const childParts: string[] = [];
    const attrParts: string[] = [];

    for (const child of this.root) {
      if (child instanceof BaseXmlComponent) {
        const s = child.toXml(context);
        if (s) childParts.push(s);
        // NextAttributeComponent.toXml() returns "" — extract attrs via prepForXml()
        if (!s && child.constructor.name === "NextAttributeComponent") {
          const prepped = child.prepForXml(context);
          if (prepped && typeof prepped === "object" && "_attr" in prepped) {
            const a = (prepped as { _attr: Record<string, unknown> })._attr;
            for (const key of Object.keys(a)) {
              attrParts.push(`${key}="${escapeXml(String(a[key]))}"`);
            }
          }
        }
      } else if (typeof child === "string") {
        childParts.push(escapeXml(child));
      } else if (child != null && typeof child === "object") {
        if ("_attr" in child) {
          const a = (child as { _attr: Record<string, unknown> })._attr;
          for (const key of Object.keys(a)) {
            attrParts.push(`${key}="${escapeXml(String(a[key]))}"`);
          }
        } else if ("_attributes" in child) {
          const a = (child as { _attributes: Record<string, unknown> })._attributes;
          for (const key of Object.keys(a)) {
            attrParts.push(`${key}="${escapeXml(String(a[key]))}"`);
          }
        } else {
          childParts.push(xml(child as Record<string, any>));
        }
      }
    }

    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    const body = childParts.join("");

    return body.length === 0
      ? `<${this.rootKey}${attrStr}/>`
      : `<${this.rootKey}${attrStr}>${body}</${this.rootKey}>`;
  }

  /**
   * @deprecated Internal use only.
   */
  public addChildElement(child: BaseXmlComponent | string): XmlComponent {
    this.root.push(child);
    return this;
  }
}

/**
 * XML component that is excluded from output if it has no meaningful content.
 */
export abstract class IgnoreIfEmptyXmlComponent extends XmlComponent {
  private readonly includeIfEmpty: boolean | undefined;

  public constructor(rootKey: string, includeIfEmpty?: boolean) {
    super(rootKey);
    this.includeIfEmpty = includeIfEmpty;
  }

  public prepForXml(context: Context): IXmlableObject | undefined {
    const result = super.prepForXml(context);

    if (this.includeIfEmpty) {
      return result;
    }
    if (
      result &&
      (typeof result[this.rootKey] !== "object" || Object.keys(result[this.rootKey]).length)
    ) {
      return result;
    }

    return undefined;
  }

  /**
   * Suppress empty elements — super.toXml() produces a self-closing tag
   * (`<tag/>`) when there are no child elements; IgnoreIfEmpty returns ""
   * in that case. Child components that already override toXml() (e.g.
   * Worksheet) are unaffected.
   */
  public override toXml(context: Context): string {
    if (this.includeIfEmpty) {
      return super.toXml(context);
    }
    const result = super.toXml(context);
    return result.endsWith("/>") ? "" : result;
  }
}
