import { ImportedXmlComponent } from "@file/xml-components";

const TABLE_STYLES_XML = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;

export class TableStyles extends ImportedXmlComponent {
    private static instance = ImportedXmlComponent.fromXmlString(TABLE_STYLES_XML);

    public constructor() {
        super("a:tblStyleLst");
    }

    public prepForXml() {
        return TableStyles.instance.prepForXml({ stack: [] });
    }
}
