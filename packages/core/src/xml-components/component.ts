/**
 * Core XML Component classes for building OOXML element trees.
 *
 * @module
 */
import { BaseXmlComponent } from "./base";
import type { IContext, IXmlableObject } from "./base";

/**
 * Empty object singleton used for empty XML elements.
 *
 * @internal
 */
export const EMPTY_OBJECT = Object.seal({});

/**
 * Base class for all XML components in OOXML documents.
 */
export abstract class XmlComponent extends BaseXmlComponent {
    /**
     * Array of child components, text nodes, and attributes.
     */
    public root: (BaseXmlComponent | IXmlableObject | string)[];

    public constructor(rootKey: string) {
        super(rootKey);
        this.root = new Array<BaseXmlComponent | string>();
    }

    /**
     * Prepares this component and its children for XML serialization.
     */
    public prepForXml(context: IContext): IXmlableObject | undefined {
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

    public prepForXml(context: IContext): IXmlableObject | undefined {
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
}
