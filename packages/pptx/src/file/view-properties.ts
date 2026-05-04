import { ImportedXmlComponent } from "@file/xml-components";

const VIEW_PROPS_XML = `<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr>
    <p:restoredLeft sz="14996" autoAdjust="0"/>
    <p:restoredTop sz="94660"/>
  </p:normalViewPr>
  <p:slideViewPr>
    <p:cSldViewPr snapToGrid="0">
      <p:cViewPr varScale="1">
        <p:scale><a:sx n="90" d="100"/><a:sy n="90" d="100"/></p:scale>
        <p:origin x="1200" y="72"/>
      </p:cViewPr>
      <p:guideLst/>
    </p:cSldViewPr>
  </p:slideViewPr>
  <p:notesTextViewPr>
    <p:cViewPr>
      <p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale>
      <p:origin x="0" y="0"/>
    </p:cViewPr>
  </p:notesTextViewPr>
  <p:gridSpacing cx="72008" cy="72008"/>
</p:viewPr>`;

export class ViewProperties extends ImportedXmlComponent {
    private static instance = ImportedXmlComponent.fromXmlString(VIEW_PROPS_XML);

    public constructor() {
        super("p:viewPr");
    }

    public prepForXml() {
        return ViewProperties.instance.prepForXml({ stack: [] });
    }
}
