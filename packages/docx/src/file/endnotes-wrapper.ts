import type { ViewWrapper } from "./document-wrapper";
import { Endnotes } from "./endnotes/endnotes";
import { Relationships } from "./relationships";

export class EndnotesWrapper implements ViewWrapper {
  private readonly endnotes: Endnotes;
  public readonly relationships: Relationships;

  public constructor() {
    this.endnotes = new Endnotes();
    this.relationships = new Relationships();
  }

  public get view(): Endnotes {
    return this.endnotes;
  }
}
