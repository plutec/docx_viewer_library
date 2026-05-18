import type { DocxDocumentModel, InlineNode, RevisionNode } from "./model.js";

export function renderDocumentAsHtml(documentModel: DocxDocumentModel): string {
  const body = documentModel.paragraphs.map(renderParagraph).join("");
  return `<article class="docx-viewer">${body}</article>`;
}

function renderParagraph(paragraph: DocxDocumentModel["paragraphs"][number]): string {
  const tagName = paragraphTagName(paragraph.styleId, paragraph.styleName);
  const attrs = [
    paragraph.styleId ? `data-paragraph-style-id="${escapeHtml(paragraph.styleId)}"` : "",
    paragraph.styleName ? `data-paragraph-style-name="${escapeHtml(paragraph.styleName)}"` : ""
  ]
    .filter(Boolean)
    .join(" ");
  const openTag = attrs ? `<${tagName} ${attrs}>` : `<${tagName}>`;
  return `${openTag}${paragraph.children.map(renderInlineNode).join("")}</${tagName}>`;
}

function renderInlineNode(node: InlineNode): string {
  if (node.type === "text") {
    return escapeHtml(node.value).replace(/\n/g, "<br/>").replace(/\t/g, "&emsp;");
  }

  return renderRevision(node);
}

function renderRevision(node: RevisionNode): string {
  const attrs = [
    `data-revision-kind="${escapeHtml(node.revision.kind)}"`,
    optionalDataAttribute("data-revision-id", node.revision.id),
    optionalDataAttribute("data-revision-author", node.revision.author),
    optionalDataAttribute("data-revision-date", node.revision.date)
  ]
    .filter(Boolean)
    .join(" ");

  const content = node.children.map(renderInlineNode).join("");

  if (node.revision.kind === "deletion" || node.revision.kind === "moveFrom") {
    return `<del ${attrs}>${content}</del>`;
  }

  return `<ins ${attrs}>${content}</ins>`;
}

function optionalDataAttribute(name: string, value: string | undefined): string {
  return value ? `${name}="${escapeHtml(value)}"` : "";
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function paragraphTagName(styleId: string | undefined, styleName: string | undefined): string {
  const key = `${styleId ?? ""} ${styleName ?? ""}`.toLowerCase();

  if (key.includes("title") || key.includes("titulo")) {
    return "h1";
  }
  if (key.includes("subtitle") || key.includes("subtitulo")) {
    return "h2";
  }
  for (const level of ["1", "2", "3", "4", "5", "6"]) {
    if (key.includes(`heading ${level}`) || key.includes(`titulo ${level}`)) {
      return `h${level}`;
    }
  }

  return "p";
}
