import { writeFileSync } from "node:fs";
// Simple example to add text to a document

import { AlignmentType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            alignment: AlignmentType.THAI_DISTRIBUTE,
            children: [
              {
                text: "บริษัทฯ มีเงินสด 41,985.00 บาท และ 25,855.66 บาทตามลำดับ เงินสดทั้งจำนวนอยู่ในความดูแลและรับผิดชอบของกรรมการ บริษัทฯบันทึกการรับชำระเงินและการจ่ายชำระเงินผ่านบัญชีเงินสดเพียงเท่านั้น ซึ่งอาจกระทบต่อความถูกต้องครบถ้วนของการบันทึกบัญชี ทั้งนี้ขึ้นอยู่กับระบบการควบคุมภายในของบริษัท",
                size: 14,
              },
            ],
          },
        },
      ],
      properties: {
        page: {
          margin: {
            bottom: "24mm",
            left: "24mm",
            right: "24mm",
            top: 0,
          },
        },
      },
    },
  ],
  styles: {
    paragraphStyles: [
      {
        basedOn: "Normal",
        id: "test",
        name: "Test",
        next: "Normal",
        paragraph: {
          indent: { left: "6.4mm" },
        },
      },
    ],
  },
});
writeFileSync("My Document.docx", buffer);
