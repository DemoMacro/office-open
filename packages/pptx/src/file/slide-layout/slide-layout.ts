import { ImportedXmlComponent } from "@file/xml-components";

const SLIDE_LAYOUT_XML = `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Slide Number Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm anchor="t"><a:off x="457200" y="666750"/><a:ext cx="914400" cy="365760"/></a:xfrm>
          <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld type="slidNum"><a:ptv/><a:endParaRPr lang="en-US"/></a:fld></a:p></a:txBody>
          <p:extLst><p:ext xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:modId="{A8390B1C-6C38-4F8F-A66D-48FD5329C7A0}"/></p:extLst>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="2"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm anchor="t"><a:off x="457200" y="2743200"/><a:ext cx="4572000" cy="365760"/></a:xfrm>
          <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld type="datetime"><a:ptv/><a:endParaRPr lang="en-US"/></a:fld></a:p></a:txBody>
          <p:extLst><p:ext xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:modId="{A8390B1C-6C38-4F8F-A66D-48FD5329C7A0}"/></p:extLst>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Footer Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" idx="3"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm anchor="t"><a:off x="457200" y="6629400"/><a:ext cx="6858000" cy="365760"/></a:xfrm>
          <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody>
          <p:extLst><p:ext xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:modId="{A8390B1C-6C38-4F8F-A66D-48FD5329C7A0}"/></p:extLst>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`;

export class DefaultSlideLayout extends ImportedXmlComponent {
    private static instance = ImportedXmlComponent.fromXmlString(SLIDE_LAYOUT_XML);

    public constructor() {
        super("p:sldLayout");
    }

    public prepForXml() {
        return DefaultSlideLayout.instance.prepForXml({ stack: [] });
    }
}
