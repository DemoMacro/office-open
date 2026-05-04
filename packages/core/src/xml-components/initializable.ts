/**
 * Initializable XML Component module.
 *
 * @module
 */
import { XmlComponent } from "./component";

/**
 * XML component that can be initialized from another component.
 */
export abstract class InitializableXmlComponent extends XmlComponent {
    public constructor(rootKey: string, initComponent?: XmlComponent) {
        super(rootKey);

        if (initComponent instanceof XmlComponent) {
            this.root = initComponent.root;
        }
    }
}
