// Simple example to add text to a document

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    Packer,
    Paragraph,
    TextRun,
    convertMillimetersToTwip,
} from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    alignment: AlignmentType.THAI_DISTRIBUTE,
                    children: [
                        new TextRun({
                            text: "บริษัทฯ มีเงินสด 41,985.00 บาท และ 25,855.66 บาทตามลำดับ เงินสดทั้งจำนวนอยู่ในความดูแลและรับผิดชอบของกรรมการ บริษัทฯบันทึกการรับชำระเงินและการจ่ายชำระเงินผ่านบัญชีเงินสดเพียงเท่านั้น ซึ่งอาจกระทบต่อความถูกต้องครบถ้วนของการบันทึกบัญชี ทั้งนี้ขึ้นอยู่กับระบบการควบคุมภายในของบริษัท",
                            size: 28,
                        }),
                    ],
                }),
            ],
            properties: {
                page: {
                    margin: {
                        bottom: convertMillimetersToTwip(24),
                        left: convertMillimetersToTwip(24),
                        right: convertMillimetersToTwip(24),
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
                    indent: { left: convertMillimetersToTwip(6.4) },
                },
            },
        ],
    },
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
