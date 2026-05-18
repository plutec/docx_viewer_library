import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type { DocxDocumentModel, InlineNode, ParagraphNode, RevisionMetadata } from "./model.js";

type XmlNode = Record<string, unknown>;
type StyleMap = Map<string, string>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  preserveOrder: true,
  trimValues: false
});

export async function parseDocx(input: ArrayBuffer | Uint8Array): Promise<DocxDocumentModel> {
  const zip = await JSZip.loadAsync(input);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  const stylesXml = await zip.file("word/styles.xml")?.async("string");

  if (!documentXml) {
    throw new Error("The DOCX file does not contain word/document.xml");
  }

  const xml = parser.parse(documentXml) as unknown[];
  const styles = stylesXml ? extractParagraphStyles(parser.parse(stylesXml) as unknown[]) : new Map();
  const paragraphs = extractParagraphs(xml, styles);

  return {
    type: "document",
    paragraphs
  };
}

function extractParagraphs(nodes: unknown[], styles: StyleMap): ParagraphNode[] {
  const body = findChildren(nodes, "w:document")
    .flatMap((node) => findChildren(asNodeArray(node), "w:body"));

  return body.flatMap((bodyNode) =>
    findChildren(asNodeArray(bodyNode), "w:p").map((paragraphNode) =>
      buildParagraph(asNodeArray(paragraphNode), styles)
    )
  );
}

function buildParagraph(nodes: unknown[], styles: StyleMap): ParagraphNode {
  const styleId = extractParagraphStyleId(nodes);

  return {
    type: "paragraph",
    styleId,
    styleName: styleId ? styles.get(styleId) : undefined,
    children: extractInlineNodes(nodes)
  };
}

function extractInlineNodes(nodes: unknown[]): InlineNode[] {
  const results: InlineNode[] = [];

  for (const entry of nodes) {
    if (!isXmlNode(entry)) {
      continue;
    }

    const [name, value] = firstEntry(entry);

    if (name === "w:r") {
      results.push(...extractRunChildren(asNodeArray(value)));
      continue;
    }

    if (name === "w:ins" || name === "w:del" || name === "w:moveFrom" || name === "w:moveTo") {
      results.push({
        type: "revision",
        revision: revisionFromNode(name, asXmlNode(value)),
        children: extractInlineNodes(asNodeArray(value))
      });
    }
  }

  return mergeAdjacentTextNodes(results);
}

function extractRunChildren(nodes: unknown[]): InlineNode[] {
  const output: InlineNode[] = [];

  for (const entry of nodes) {
    if (!isXmlNode(entry)) {
      continue;
    }

    const [name, value] = firstEntry(entry);

    if (name === "w:t" || name === "w:delText") {
      const text = extractTextValue(value);
      if (text) {
        output.push({ type: "text", value: text });
      }
      continue;
    }

    if (name === "w:tab") {
      output.push({ type: "text", value: "\t" });
      continue;
    }

    if (name === "w:br") {
      output.push({ type: "text", value: "\n" });
    }
  }

  return output;
}

function revisionFromNode(name: string, node: XmlNode): RevisionMetadata {
  const kindMap: Record<string, RevisionMetadata["kind"]> = {
    "w:ins": "insertion",
    "w:del": "deletion",
    "w:moveFrom": "moveFrom",
    "w:moveTo": "moveTo"
  };

  return {
    kind: kindMap[name] ?? "formatChange",
    id: getAttribute(node, "@_w:id"),
    author: getAttribute(node, "@_w:author"),
    date: getAttribute(node, "@_w:date")
  };
}

function extractParagraphStyleId(nodes: unknown[]): string | undefined {
  const properties = findChildren(nodes, "w:pPr");
  for (const propertyNode of properties) {
    for (const entry of asNodeArray(propertyNode)) {
      if (!isXmlNode(entry)) {
        continue;
      }
      const [name, value] = firstEntry(entry);
      if (name === "w:pStyle") {
        return getAttribute(asXmlNode(value), "@_w:val");
      }
    }
  }
  return undefined;
}

function extractParagraphStyles(nodes: unknown[]): StyleMap {
  const styles = new Map<string, string>();
  const styleNodes = findChildren(nodes, "w:styles").flatMap((node) => findChildren(asNodeArray(node), "w:style"));

  for (const styleNode of styleNodes) {
    const entries = asNodeArray(styleNode);
    const attrs = collectAttributes(entries);
    if (attrs["@_w:type"] !== "paragraph" || typeof attrs["@_w:styleId"] !== "string") {
      continue;
    }

    const styleId = attrs["@_w:styleId"];
    const styleName = entries
      .filter(isXmlNode)
      .map(firstEntry)
      .find(([name]) => name === "w:name");

    const resolvedName = styleName ? getAttribute(asXmlNode(styleName[1]), "@_w:val") : undefined;
    styles.set(styleId, resolvedName ?? styleId);
  }

  return styles;
}

function extractTextValue(value: unknown): string {
  if (typeof value === "string") {
    return value;
  }

  if (Array.isArray(value)) {
    return value
      .map((item) => {
        if (typeof item === "string") {
          return item;
        }
        if (isXmlNode(item) && typeof item["#text"] === "string") {
          return item["#text"];
        }
        return "";
      })
      .join("");
  }

  if (isXmlNode(value) && typeof value["#text"] === "string") {
    return value["#text"];
  }

  return "";
}

function getAttribute(node: XmlNode, attributeName: string): string | undefined {
  const value = node[attributeName];
  return typeof value === "string" ? value : undefined;
}

function findChildren(nodes: unknown[], targetName: string): unknown[] {
  return nodes
    .filter(isXmlNode)
    .flatMap((node) => {
      const [name, value] = firstEntry(node);
      return name === targetName ? [value] : [];
    });
}

function collectAttributes(nodes: unknown[]): XmlNode {
  const attributes: XmlNode = {};
  for (const entry of nodes) {
    if (!isXmlNode(entry)) {
      continue;
    }
    for (const [name, value] of Object.entries(entry)) {
      if (name.startsWith("@_")) {
        attributes[name] = value;
      }
    }
  }
  return attributes;
}

function mergeAdjacentTextNodes(nodes: InlineNode[]): InlineNode[] {
  const merged: InlineNode[] = [];

  for (const node of nodes) {
    const previous = merged.at(-1);

    if (node.type === "text" && previous?.type === "text") {
      previous.value += node.value;
      continue;
    }

    merged.push(node);
  }

  return merged;
}

function asNodeArray(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function asXmlNode(value: unknown): XmlNode {
  return isXmlNode(value) ? value : {};
}

function isXmlNode(value: unknown): value is XmlNode {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function firstEntry(node: XmlNode): [string, unknown] {
  const entry = Object.entries(node)[0];
  if (!entry) {
    return ["", undefined];
  }
  return entry;
}
