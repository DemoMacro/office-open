import { buildAttrObject, BuilderElement, XmlComponent } from "@file/xml-components";

export interface CommentEntry {
  readonly authorId: number;
  readonly idx: number;
  readonly date?: string;
  readonly x: number;
  readonly y: number;
  readonly text: string;
}

/**
 * p:cmLst — Comment list stored in `ppt/comments/slideN.xml`.
 */
export class SlideCommentList extends XmlComponent {
  public constructor(comments: readonly CommentEntry[]) {
    super("p:cmLst");
    this.root.push(
      buildAttrObject({
        "xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      }),
    );
    for (const comment of comments) {
      this.root.push(
        new BuilderElement({
          name: "p:cm",
          attributes: {
            authorId: { key: "authorId", value: comment.authorId },
            idx: { key: "idx", value: comment.idx },
            ...(comment.date != null ? { dt: { key: "dt", value: comment.date } } : {}),
          },
          children: [
            new BuilderElement({
              name: "p:pos",
              attributes: {
                x: { key: "x", value: comment.x },
                y: { key: "y", value: comment.y },
              },
            }),
            new BuilderElement({
              name: "p:text",
              children: [comment.text],
            }),
          ],
        }),
      );
    }
  }
}
