import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

export interface AuthorEntry {
    readonly id: number;
    readonly name: string;
    readonly initials: string;
    readonly clrIdx: number;
    readonly lastIdx: number;
}

/**
 * p:cmAuthorLst — Comment author list stored in `ppt/commentAuthors.xml`.
 */
export class CommentAuthorList extends XmlComponent {
    public constructor(authors: readonly AuthorEntry[]) {
        super("p:cmAuthorLst");
        this.root.push(
            new NextAttributeComponent({
                "xmlns:p": {
                    key: "xmlns:p",
                    value: "http://schemas.openxmlformats.org/presentationml/2006/main",
                },
            }),
        );
        for (const author of authors) {
            this.root.push(
                new BuilderElement({
                    name: "p:cmAuthor",
                    attributes: {
                        id: { key: "id", value: author.id },
                        name: { key: "name", value: author.name },
                        initials: { key: "initials", value: author.initials },
                        clrIdx: { key: "clrIdx", value: author.clrIdx },
                        lastIdx: { key: "lastIdx", value: author.lastIdx },
                    },
                }),
            );
        }
    }
}
