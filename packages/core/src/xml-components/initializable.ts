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
    public constructor(rootKey: string, initComponent?: InitializableXmlComponent) {
        super(rootKey);

        if (initComponent) {
            this.root = initComponent.root;
        }
    }
}
