export type RevisionKind = "insertion" | "deletion" | "moveFrom" | "moveTo" | "formatChange";

export interface RevisionMetadata {
  id?: string;
  author?: string;
  date?: string;
  kind: RevisionKind;
}

export interface TextNode {
  type: "text";
  value: string;
}

export interface RevisionNode {
  type: "revision";
  revision: RevisionMetadata;
  children: InlineNode[];
}

export type InlineNode = TextNode | RevisionNode;

export interface ParagraphNode {
  type: "paragraph";
  styleId?: string;
  styleName?: string;
  children: InlineNode[];
}

export interface DocxDocumentModel {
  type: "document";
  paragraphs: ParagraphNode[];
}
