// Simple example to add text to a document

import { writeFileSync } from "node:fs";

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
                size: 18,
                text: "บริษัท บิสกิด จำกัด (บริษัทฯ) ได้จดทะเบียนจัดตั้งขึ้นเป็นบริษัทจำกัดตามประมวลกฎหมายแพ่งและพาณิชย์ของประเทศไทย เมื่อวันที่ 30 พฤษภาคม 2561 ทะเบียนนิติบุคคลเลขที่ 0845561005665",
              },
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
