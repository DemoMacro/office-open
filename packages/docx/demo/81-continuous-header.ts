// Example of a continuous header

import { writeFileSync } from "node:fs";

import { SectionType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  creator: "Creator",
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer ac suscipit orci, in lobortis risus. Nulla vehicula rutrum finibus. Nullam consequat, magna in vehicula commodo, enim massa consectetur nisl, sit amet rutrum nunc ante vel lorem. Sed sit amet scelerisque velit. Proin non quam eget mauris aliquet posuere a sed orci. Proin posuere ante suscipit neque dignissim hendrerit. Pellentesque eget dapibus metus. Donec at mollis mauris. Vestibulum sit amet scelerisque nulla. Vivamus ipsum erat, tempor sed volutpat non, molestie at odio. Vivamus lectus ligula, finibus at mattis vitae, euismod sed tellus. Etiam neque massa, faucibus a fringilla nec, mollis at ex. Aliquam eget nibh tortor. Sed ut viverra libero. Nulla facilisis bibendum quam eget porttitor.",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Sed eget nunc ac turpis facilisis volutpat. Duis eget arcu vitae neque porta hendrerit. Proin vel ante nulla. Duis congue efficitur dui. Suspendisse potenti. Aliquam aliquam nibh eu ipsum sagittis efficitur. Quisque sagittis metus dui, vitae suscipit tortor sollicitudin at. Suspendisse convallis, sem ac ornare condimentum, odio ipsum dapibus justo, a aliquam risus massa ut enim. Mauris vel placerat nibh. Ut iaculis vitae nibh at elementum. Quisque hendrerit et magna vitae mollis. Duis dictum euismod leo, at cursus risus sodales sed.",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Sed gravida commodo felis, at aliquet risus volutpat ut. Nam nec ex eleifend tellus sodales volutpat nec ac nibh. Vestibulum pretium, leo vitae lobortis accumsan, urna libero euismod ante, consequat aliquam enim risus id nisl. Donec sagittis, justo eu luctus posuere, leo purus pellentesque turpis, eget volutpat mi leo vitae lacus. Etiam ante ante, posuere at augue non, lacinia ornare purus. Praesent vitae velit in enim congue maximus. Vivamus tincidunt fringilla neque. Curabitur fermentum justo nec sapien porttitor, ac ullamcorper nisi imperdiet. Orci varius natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Sed non orci vel eros egestas eleifend sit amet a diam. Duis mattis at ligula quis faucibus. Donec elementum lacus velit, a vehicula nunc gravida a. Phasellus eget nunc vehicula, varius velit a, maximus velit. Sed a suscipit nisi, non hendrerit felis. Proin mattis facilisis massa, quis elementum neque fringilla non.",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Sed gravida commodo felis, at aliquet risus volutpat ut. Nam nec ex eleifend tellus sodales volutpat nec ac nibh. Vestibulum pretium, leo vitae lobortis accumsan, urna libero euismod ante, consequat aliquam enim risus id nisl. Donec sagittis, justo eu luctus posuere, leo purus pellentesque turpis, eget volutpat mi leo vitae lacus. Etiam ante ante, posuere at augue non, lacinia ornare purus. Praesent vitae velit in enim congue maximus. Vivamus tincidunt fringilla neque. Curabitur fermentum justo nec sapien porttitor, ac ullamcorper nisi imperdiet. Orci varius natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Sed non orci vel eros egestas eleifend sit amet a diam. Duis mattis at ligula quis faucibus. Donec elementum lacus velit, a vehicula nunc gravida a. Phasellus eget nunc vehicula, varius velit a, maximus velit. Sed a suscipit nisi, non hendrerit felis. Proin mattis facilisis massa, quis elementum neque fringilla non.",
              },
            ],
          },
        },
        {
          paragraph: {
            spacing: {
              after: 500,
            },
            children: [
              {
                text: "The first section ends after this paragraph.",
              },
            ],
          },
        },
      ],
      footers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  text: "FOOTER PAGE TWO AND FOLLOWING PAGES",
                },
              ],
            },
          },
        ],
        first: [
          {
            paragraph: {
              children: [
                {
                  text: "FOOTER PAGE ONE",
                },
              ],
            },
          },
        ],
      },
      headers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  text: "HEADER PAGE TWO AND FOLLOWING PAGES",
                },
              ],
            },
          },
        ],
        first: [
          {
            paragraph: {
              children: [
                {
                  text: "HEADER PAGE ONE",
                },
              ],
            },
          },
        ],
      },
      properties: { titlePage: true },
    },
    {
      children: [
        {
          paragraph: {
            children: [
              {
                text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer ac suscipit orci, in lobortis risus. Nulla vehicula rutrum finibus. Nullam consequat, magna in vehicula commodo, enim massa consectetur nisl, sit amet rutrum nunc ante vel lorem. Sed sit amet scelerisque velit. Proin non quam eget mauris aliquet posuere a sed orci. Proin posuere ante suscipit neque dignissim hendrerit. Pellentesque eget dapibus metus. Donec at mollis mauris. Vestibulum sit amet scelerisque nulla. Vivamus ipsum erat, tempor sed volutpat non, molestie at odio. Vivamus lectus ligula, finibus at mattis vitae, euismod sed tellus. Etiam neque massa, faucibus a fringilla nec, mollis at ex. Aliquam eget nibh tortor. Sed ut viverra libero. Nulla facilisis bibendum quam eget porttitor.",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Sed eget nunc ac turpis facilisis volutpat. Duis eget arcu vitae neque porta hendrerit. Proin vel ante nulla. Duis congue efficitur dui. Suspendisse potenti. Aliquam aliquam nibh eu ipsum sagittis efficitur. Quisque sagittis metus dui, vitae suscipit tortor sollicitudin at. Suspendisse convallis, sem ac ornare condimentum, odio ipsum dapibus justo, a aliquam risus massa ut enim. Mauris vel placerat nibh. Ut iaculis vitae nibh at elementum. Quisque hendrerit et magna vitae mollis. Duis dictum euismod leo, at cursus risus sodales sed.",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Sed gravida commodo felis, at aliquet risus volutpat ut. Nam nec ex eleifend tellus sodales volutpat nec ac nibh. Vestibulum pretium, leo vitae lobortis accumsan, urna libero euismod ante, consequat aliquam enim risus id nisl. Donec sagittis, justo eu luctus posuere, leo purus pellentesque turpis, eget volutpat mi leo vitae lacus. Etiam ante ante, posuere at augue non, lacinia ornare purus. Praesent vitae velit in enim congue maximus. Vivamus tincidunt fringilla neque. Curabitur fermentum justo nec sapien porttitor, ac ullamcorper nisi imperdiet. Orci varius natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Sed non orci vel eros egestas eleifend sit amet a diam. Duis mattis at ligula quis faucibus. Donec elementum lacus velit, a vehicula nunc gravida a. Phasellus eget nunc vehicula, varius velit a, maximus velit. Sed a suscipit nisi, non hendrerit felis. Proin mattis facilisis massa, quis elementum neque fringilla non.",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Sed gravida commodo felis, at aliquet risus volutpat ut. Nam nec ex eleifend tellus sodales volutpat nec ac nibh. Vestibulum pretium, leo vitae lobortis accumsan, urna libero euismod ante, consequat aliquam enim risus id nisl. Donec sagittis, justo eu luctus posuere, leo purus pellentesque turpis, eget volutpat mi leo vitae lacus. Etiam ante ante, posuere at augue non, lacinia ornare purus. Praesent vitae velit in enim congue maximus. Vivamus tincidunt fringilla neque. Curabitur fermentum justo nec sapien porttitor, ac ullamcorper nisi imperdiet. Orci varius natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Sed non orci vel eros egestas eleifend sit amet a diam. Duis mattis at ligula quis faucibus. Donec elementum lacus velit, a vehicula nunc gravida a. Phasellus eget nunc vehicula, varius velit a, maximus velit. Sed a suscipit nisi, non hendrerit felis. Proin mattis facilisis massa, quis elementum neque fringilla non.",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "The second section starts with the headline above. Move cursor to the end of this text and press enter until next page is generated in continuous section break mode.",
              },
            ],
          },
        },
      ],
      properties: {
        type: SectionType.CONTINUOUS,
      },
    },
  ],
  title: "Title",
});
writeFileSync("My Document.docx", buffer);
