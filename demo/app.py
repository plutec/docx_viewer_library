from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from html import escape
from io import BytesIO
from typing import Any
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates


app = FastAPI(title="DOCX Viewer Demo")
templates = Jinja2Templates(directory="demo/templates")

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "html_output": None,
            "error": None,
            "document_info": None,
        },
    )


@app.post("/", response_class=HTMLResponse)
async def upload_docx(request: Request, file: UploadFile = File(...)) -> HTMLResponse:
    html_output: str | None = None
    error: str | None = None
    document_info: dict[str, Any] | None = None

    try:
        filename = file.filename or "uploaded.docx"
        content = await file.read()
        document = parse_docx(content)
        html_output = render_document_html(document)
        document_info = {
            "filename": filename,
            "block_count": len(document["blocks"]),
            "paragraph_count": count_paragraphs(document["blocks"]),
            "table_count": count_tables(document["blocks"]),
            "revision_count": count_revisions(document["blocks"]),
            "processed_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        }
    except Exception as exc:
        error = str(exc)

    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "html_output": html_output,
            "error": error,
            "document_info": document_info,
        },
    )


def parse_docx(content: bytes) -> dict[str, Any]:
    with ZipFile(BytesIO(content)) as archive:
        try:
            document_xml = archive.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("The file does not contain word/document.xml") from exc
        try:
            styles_xml = archive.read("word/styles.xml")
        except KeyError:
            styles_xml = None

    root = ET.fromstring(document_xml)
    styles = parse_styles(styles_xml) if styles_xml else {"paragraph": {}, "character": {}}
    body = root.find(".//w:body", NS)
    blocks = parse_body_blocks(body, styles) if body is not None else []

    return {
        "type": "document",
        "blocks": blocks,
    }


def parse_body_blocks(body: ET.Element, styles: dict[str, dict[str, dict[str, Any]]]) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []

    for child in list(body):
        tag = local_name(child.tag)
        if tag == "p":
            blocks.append(parse_paragraph(child, styles))
        elif tag == "tbl":
            blocks.append(parse_table(child, styles))

    return blocks


def parse_table(table: ET.Element, styles: dict[str, dict[str, dict[str, Any]]]) -> dict[str, Any]:
    props = parse_table_props(table.find("./w:tblPr", NS))
    props["gridColsPt"] = parse_table_grid(table.find("./w:tblGrid", NS))
    rows = [parse_table_row(row, styles) for row in table.findall("./w:tr", NS)]
    return {
        "type": "table",
        "props": props,
        "rows": rows,
    }


def parse_table_row(row: ET.Element, styles: dict[str, dict[str, dict[str, Any]]]) -> dict[str, Any]:
    props = parse_table_row_props(row.find("./w:trPr", NS))
    cells = [parse_table_cell(cell, styles) for cell in row.findall("./w:tc", NS)]
    return {
        "type": "tableRow",
        "props": props,
        "cells": cells,
    }


def parse_table_cell(cell: ET.Element, styles: dict[str, dict[str, dict[str, Any]]]) -> dict[str, Any]:
    tc_pr = cell.find("./w:tcPr", NS)
    props = parse_table_cell_props(tc_pr)
    blocks: list[dict[str, Any]] = []
    for child in list(cell):
        tag = local_name(child.tag)
        if tag == "p":
            blocks.append(parse_paragraph(child, styles))
        elif tag == "tbl":
            blocks.append(parse_table(child, styles))

    return {
        "type": "tableCell",
        "props": props,
        "blocks": blocks,
    }


def parse_table_grid(node: ET.Element | None) -> list[float]:
    if node is None:
        return []

    widths: list[float] = []
    for grid_col in node.findall("./w:gridCol", NS):
        width_pt = twips_to_pt(attr(grid_col, "w"))
        if width_pt is not None:
            widths.append(width_pt)

    return widths


def parse_paragraph(paragraph: ET.Element, styles: dict[str, dict[str, dict[str, Any]]]) -> dict[str, Any]:
    style_id = paragraph_style_id(paragraph)
    style_def = resolve_style(styles["paragraph"], style_id)
    paragraph_props = merge_dicts(style_def.get("paragraph", {}), parse_paragraph_props(paragraph.find("./w:pPr", NS)))
    run_defaults = merge_dicts(style_def.get("run", {}), {})

    return {
        "type": "paragraph",
        "style_id": style_id,
        "style_name": style_def.get("name"),
        "props": paragraph_props,
        "children": parse_inline_nodes(list(paragraph), styles, run_defaults),
    }


def parse_styles(styles_xml: bytes) -> dict[str, dict[str, dict[str, Any]]]:
    root = ET.fromstring(styles_xml)
    paragraph_styles: dict[str, dict[str, Any]] = {}
    character_styles: dict[str, dict[str, Any]] = {}

    for style in root.findall("./w:style", NS):
        style_id = style.attrib.get(qname("styleId"))
        style_type = style.attrib.get(qname("type"))
        if not style_id or style_type not in {"paragraph", "character"}:
            continue

        name_node = style.find("./w:name", NS)
        style_data = {
            "name": name_node.attrib.get(qname("val"), style_id) if name_node is not None else style_id,
            "based_on": attr(style.find("./w:basedOn", NS), "val"),
            "paragraph": parse_paragraph_props(style.find("./w:pPr", NS)),
            "run": parse_run_props(style.find("./w:rPr", NS)),
        }

        if style_type == "paragraph":
            paragraph_styles[style_id] = style_data
        else:
            character_styles[style_id] = style_data

    return {
        "paragraph": paragraph_styles,
        "character": character_styles,
    }


def resolve_style(style_map: dict[str, dict[str, Any]], style_id: str | None) -> dict[str, Any]:
    if not style_id or style_id not in style_map:
        return {"name": None, "paragraph": {}, "run": {}}

    style = style_map[style_id]
    parent = resolve_style(style_map, style.get("based_on"))

    return {
        "name": style.get("name") or parent.get("name"),
        "paragraph": merge_dicts(parent.get("paragraph", {}), style.get("paragraph", {})),
        "run": merge_dicts(parent.get("run", {}), style.get("run", {})),
    }


def paragraph_style_id(paragraph: ET.Element) -> str | None:
    p_style = paragraph.find("./w:pPr/w:pStyle", NS)
    return attr(p_style, "val")


def parse_inline_nodes(
    elements: list[ET.Element],
    styles: dict[str, dict[str, dict[str, Any]]],
    inherited_run_props: dict[str, Any],
) -> list[dict[str, Any]]:
    nodes: list[dict[str, Any]] = []

    for element in elements:
        tag = local_name(element.tag)

        if tag == "r":
            nodes.extend(parse_run_children(element, styles, inherited_run_props))
            continue

        if tag in {"ins", "del", "moveFrom", "moveTo"}:
            nodes.append(
                {
                    "type": "revision",
                    "revision": revision_metadata(tag, element),
                    "children": parse_inline_nodes(list(element), styles, inherited_run_props),
                }
            )

    return merge_adjacent_text(nodes)


def parse_run_children(
    run: ET.Element,
    styles: dict[str, dict[str, dict[str, Any]]],
    inherited_run_props: dict[str, Any],
) -> list[dict[str, Any]]:
    nodes: list[dict[str, Any]] = []
    run_props = merge_dicts(inherited_run_props, resolve_run_props(run, styles))

    for element in list(run):
        tag = local_name(element.tag)

        if tag in {"t", "delText"}:
            text = element.text or ""
            if text:
                nodes.append({"type": "text", "value": text, "props": deepcopy(run_props)})
            continue

        if tag == "tab":
            nodes.append({"type": "text", "value": "\t", "props": deepcopy(run_props)})
            continue

        if tag == "br":
            nodes.append({"type": "text", "value": "\n", "props": deepcopy(run_props)})

    return nodes


def resolve_run_props(run: ET.Element, styles: dict[str, dict[str, dict[str, Any]]]) -> dict[str, Any]:
    run_pr = run.find("./w:rPr", NS)
    direct_props = parse_run_props(run_pr)
    style_id = attr(run_pr.find("./w:rStyle", NS), "val") if run_pr is not None else None
    style_props = resolve_style(styles["character"], style_id).get("run", {})
    return merge_dicts(style_props, direct_props)


def parse_paragraph_props(node: ET.Element | None) -> dict[str, Any]:
    if node is None:
        return {}

    props: dict[str, Any] = {}

    jc = attr(node.find("./w:jc", NS), "val")
    if jc:
        props["align"] = jc

    spacing = node.find("./w:spacing", NS)
    if spacing is not None:
        before = twips_to_pt(attr(spacing, "before"))
        after = twips_to_pt(attr(spacing, "after"))
        line = twips_to_pt(attr(spacing, "line"))
        if before is not None:
            props["marginTopPt"] = before
        if after is not None:
            props["marginBottomPt"] = after
        if line is not None:
            props["lineHeightPt"] = line

    ind = node.find("./w:ind", NS)
    if ind is not None:
        left = twips_to_pt(attr(ind, "left"))
        right = twips_to_pt(attr(ind, "right"))
        first_line = twips_to_pt(attr(ind, "firstLine"))
        hanging = twips_to_pt(attr(ind, "hanging"))
        if left is not None:
            props["marginLeftPt"] = left
        if right is not None:
            props["marginRightPt"] = right
        if first_line is not None:
            props["textIndentPt"] = first_line
        if hanging is not None:
            props["textIndentPt"] = -hanging

    shading = node.find("./w:shd", NS)
    fill = resolve_fill(shading)
    if fill:
        props["backgroundColor"] = fill

    borders = node.find("./w:pBdr", NS)
    if borders is not None:
        border_css = []
        for edge in ["top", "right", "bottom", "left"]:
            edge_node = borders.find(f"./w:{edge}", NS)
            if edge_node is None:
                continue
            val = attr(edge_node, "val")
            if val in {None, "nil", "none"}:
                continue
            color = normalize_color(attr(edge_node, "color") or "000000")
            size = attr(edge_node, "sz")
            space = attr(edge_node, "space")
            px = max(1, round((int(size) / 8) if size and size.isdigit() else 1))
            border_css.append(f"border-{edge}: {px}px solid {color}")
            if space and space.isdigit():
                props[f"border{edge.capitalize()}SpacePt"] = round(int(space) * 0.75, 2)
        if border_css:
            props["borderCss"] = "; ".join(border_css)

    return props


def parse_run_props(node: ET.Element | None) -> dict[str, Any]:
    if node is None:
        return {}

    props: dict[str, Any] = {}

    if node.find("./w:b", NS) is not None:
        props["bold"] = True
    if node.find("./w:i", NS) is not None:
        props["italic"] = True
    if node.find("./w:u", NS) is not None:
        underline = attr(node.find("./w:u", NS), "val") or "single"
        if underline != "none":
            props["underline"] = underline

    color = attr(node.find("./w:color", NS), "val")
    if color and color not in {"auto", "000000"}:
        props["color"] = normalize_color(color)

    highlight = attr(node.find("./w:highlight", NS), "val")
    if highlight:
        props["highlight"] = highlight_to_css(highlight)

    shading = resolve_fill(node.find("./w:shd", NS))
    if shading:
        props["backgroundColor"] = shading

    font_size = half_points_to_pt(attr(node.find("./w:sz", NS), "val"))
    if font_size is not None:
        props["fontSizePt"] = font_size

    fonts = node.find("./w:rFonts", NS)
    if fonts is not None:
        font_family = fonts.attrib.get(qname("ascii")) or fonts.attrib.get(qname("hAnsi"))
        if font_family:
            props["fontFamily"] = font_family

    return props


def parse_table_cell_props(node: ET.Element | None) -> dict[str, Any]:
    if node is None:
        return {}

    props: dict[str, Any] = {}
    fill = resolve_fill(node.find("./w:shd", NS))
    if fill:
        props["backgroundColor"] = fill

    width = attr(node.find("./w:tcW", NS), "w")
    width_type = attr(node.find("./w:tcW", NS), "type")
    if width and width_type == "dxa":
        width_pt = twips_to_pt(width)
        if width_pt is not None:
            props["widthPt"] = width_pt

    grid_span = attr(node.find("./w:gridSpan", NS), "val")
    if grid_span:
        props["colSpan"] = int(grid_span)

    borders = node.find("./w:tcBorders", NS)
    if borders is not None:
        border_css = []
        for edge in ["top", "right", "bottom", "left"]:
            edge_node = borders.find(f"./w:{edge}", NS)
            if edge_node is None:
                continue
            val = attr(edge_node, "val")
            if val in {None, "nil", "none"}:
                continue
            color = normalize_color(attr(edge_node, "color") or "000000")
            size = attr(edge_node, "sz")
            px = max(1, round((int(size) / 8) if size and size.isdigit() else 1))
            border_css.append(f"border-{edge}: {px}px solid {color}")
        if border_css:
            props["borderCss"] = "; ".join(border_css)

    return props


def parse_table_props(node: ET.Element | None) -> dict[str, Any]:
    if node is None:
        return {}

    props: dict[str, Any] = {}

    width = attr(node.find("./w:tblW", NS), "w")
    width_type = attr(node.find("./w:tblW", NS), "type")
    if width and width_type == "dxa":
        width_pt = twips_to_pt(width)
        if width_pt is not None:
            props["widthPt"] = width_pt

    indent = attr(node.find("./w:tblInd", NS), "w")
    indent_type = attr(node.find("./w:tblInd", NS), "type")
    if indent and indent_type == "dxa":
        indent_pt = twips_to_pt(indent)
        if indent_pt is not None:
            props["marginLeftPt"] = indent_pt

    shading = resolve_fill(node.find("./w:shd", NS))
    if shading:
        props["backgroundColor"] = shading

    cell_margins = node.find("./w:tblCellMar", NS)
    if cell_margins is not None:
        for edge, key in [
            ("top", "cellPaddingTopPt"),
            ("right", "cellPaddingRightPt"),
            ("bottom", "cellPaddingBottomPt"),
            ("left", "cellPaddingLeftPt"),
        ]:
            edge_node = cell_margins.find(f"./w:{edge}", NS)
            width = attr(edge_node, "w")
            width_type = attr(edge_node, "type")
            if width and width_type == "dxa":
                padding_pt = twips_to_pt(width)
                if padding_pt is not None:
                    props[key] = padding_pt

    borders = node.find("./w:tblBorders", NS)
    if borders is not None:
        border_css = []
        for edge in ["top", "right", "bottom", "left", "insideH", "insideV"]:
            edge_node = borders.find(f"./w:{edge}", NS)
            if edge_node is None:
                continue
            val = attr(edge_node, "val")
            if val in {None, "nil", "none"}:
                continue
            color = normalize_color(attr(edge_node, "color") or "000000")
            size = attr(edge_node, "sz")
            px = max(1, round((int(size) / 8) if size and size.isdigit() else 1))
            css_edge = edge.replace("insideH", "inside-horizontal").replace("insideV", "inside-vertical")
            border_css.append(f"{css_edge}:{px}px solid {color}")
        if border_css:
            props["borderCss"] = "; ".join(border_css)

    return props


def parse_table_row_props(node: ET.Element | None) -> dict[str, Any]:
    if node is None:
        return {}

    props: dict[str, Any] = {}
    height_node = node.find("./w:trHeight", NS)
    if height_node is not None:
        height_pt = twips_to_pt(attr(height_node, "val"))
        if height_pt is not None:
            props["heightPt"] = height_pt
            props["heightRule"] = attr(height_node, "hRule") or "auto"

    return props


def revision_metadata(tag: str, element: ET.Element) -> dict[str, Any]:
    kind_map = {
        "ins": "insertion",
        "del": "deletion",
        "moveFrom": "moveFrom",
        "moveTo": "moveTo",
    }

    return {
        "kind": kind_map.get(tag, "formatChange"),
        "id": element.attrib.get(qname("id")),
        "author": element.attrib.get(qname("author")),
        "date": element.attrib.get(qname("date")),
    }


def count_tables(blocks: list[dict[str, Any]]) -> int:
    total = 0
    stack = list(blocks)
    while stack:
        block = stack.pop()
        if block["type"] == "table":
            total += 1
            for row in block["rows"]:
                for cell in row["cells"]:
                    stack.extend(cell["blocks"])
    return total


def count_paragraphs(blocks: list[dict[str, Any]]) -> int:
    total = 0
    stack = list(blocks)
    while stack:
        block = stack.pop()
        if block["type"] == "paragraph":
            total += 1
        elif block["type"] == "table":
            for row in block["rows"]:
                for cell in row["cells"]:
                    stack.extend(cell["blocks"])
    return total


def count_revisions(blocks: list[dict[str, Any]]) -> int:
    total = 0
    block_stack = list(blocks)

    while block_stack:
        block = block_stack.pop()
        if block["type"] == "paragraph":
            inline_stack = list(block["children"])
            while inline_stack:
                node = inline_stack.pop()
                if node["type"] == "revision":
                    total += 1
                    inline_stack.extend(node["children"])
        elif block["type"] == "table":
            for row in block["rows"]:
                for cell in row["cells"]:
                    block_stack.extend(cell["blocks"])

    return total


def merge_adjacent_text(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: list[dict[str, Any]] = []

    for node in nodes:
        if (
            merged
            and node["type"] == "text"
            and merged[-1]["type"] == "text"
            and merged[-1].get("props") == node.get("props")
        ):
            merged[-1]["value"] += node["value"]
            continue
        merged.append(node)

    return merged


def render_document_html(document: dict[str, Any]) -> str:
    body = "".join(render_block(block) for block in document["blocks"])
    return f'<article class="docx-viewer">{body}</article>'


def render_block(block: dict[str, Any]) -> str:
    if block["type"] == "paragraph":
        return render_paragraph(block)
    if block["type"] == "table":
        return render_table(block)
    return ""


def render_table(table: dict[str, Any]) -> str:
    attrs = [class_attr("docx-table")]
    table_props = table.get("props", {})
    style = css_from_map(table_css(table_props))
    if style:
        attrs.append(f'style="{style}"')
    colgroup = render_table_colgroup(table_props)
    rows = "".join(render_table_row(row, table_props) for row in table["rows"])
    return f"<table {' '.join(attrs)}>{colgroup}<tbody>{rows}</tbody></table>"


def render_table_colgroup(props: dict[str, Any]) -> str:
    grid_cols = props.get("gridColsPt") or []
    if not grid_cols:
        return ""

    cols = []
    for width in grid_cols:
        cols.append(f'<col style="width: {pt(width)}" />')
    return f"<colgroup>{''.join(cols)}</colgroup>"


def render_table_row(row: dict[str, Any], table_props: dict[str, Any]) -> str:
    cells = "".join(render_table_cell(cell, table_props) for cell in row["cells"])
    style = css_from_map(table_row_css(row.get("props", {})))
    if style:
        return f'<tr style="{style}">{cells}</tr>'
    return f"<tr>{cells}</tr>"


def render_table_cell(cell: dict[str, Any], table_props: dict[str, Any]) -> str:
    attrs = [class_attr("docx-cell")]
    col_span = cell["props"].get("colSpan")
    if col_span:
        attrs.append(f'colspan="{int(col_span)}"')
    style = css_from_map(cell_css(cell["props"], table_props))
    if style:
        attrs.append(f'style="{style}"')
    content = "".join(render_block(block) for block in cell["blocks"])
    return f"<td {' '.join(attrs)}>{content}</td>"


def render_paragraph(paragraph: dict[str, Any]) -> str:
    tag_name = paragraph_tag_name(paragraph.get("style_id"), paragraph.get("style_name"))
    attrs = [
        class_attr(paragraph_class_name(paragraph.get("style_id"), paragraph.get("style_name"))),
        data_attr("data-paragraph-style-id", paragraph.get("style_id")),
        data_attr("data-paragraph-style-name", paragraph.get("style_name")),
    ]
    style = css_from_map(paragraph_css(paragraph["props"]))
    if style:
        attrs.append(f'style="{style}"')
    attrs = [value for value in attrs if value]

    tabbed = split_inline_nodes_on_first_tab(paragraph["children"])
    if tabbed is not None:
        left_html = "".join(render_inline(node) for node in tabbed["left"])
        right_html = "".join(render_inline(node) for node in tabbed["right"])
        classes = [paragraph_class_name(paragraph.get("style_id"), paragraph.get("style_name")), "tabbed-paragraph"]
        open_attrs = [
            class_attr(" ".join(classes)),
            data_attr("data-paragraph-style-id", paragraph.get("style_id")),
            data_attr("data-paragraph-style-name", paragraph.get("style_name")),
        ]
        if style:
            open_attrs.append(f'style="{style}"')
        open_attrs = [value for value in open_attrs if value]
        return (
            f"<{tag_name} {' '.join(open_attrs)}>"
            f'<span class="tabbed-left">{left_html}</span>'
            f'<span class="tabbed-right">{right_html}</span>'
            f"</{tag_name}>"
        )

    children = "".join(render_inline(node) for node in paragraph["children"])
    return f"<{tag_name} {' '.join(attrs)}>{children}</{tag_name}>"


def render_inline(node: dict[str, Any]) -> str:
    if node["type"] == "text":
        text = render_text(node["value"])
        style = css_from_map(run_css(node.get("props", {})))
        if style:
            return f'<span style="{style}">{text}</span>'
        return text

    revision = node["revision"]
    children = "".join(render_inline(child) for child in node["children"])
    revision_kind = revision.get("kind", "formatChange")
    metadata = " ".join(
        value
        for value in [
            class_attr(f"revision revision-{revision_kind}"),
            data_attr("data-revision-kind", revision.get("kind")),
            data_attr("data-revision-id", revision.get("id")),
            data_attr("data-revision-author", revision.get("author")),
            data_attr("data-revision-date", revision.get("date")),
            data_attr("data-revision-label", revision_label(revision)),
        ]
        if value
    )

    if revision_kind in {"deletion", "moveFrom"}:
        return f"<del {metadata}>{children}</del>"

    return f"<ins {metadata}>{children}</ins>"


def revision_label(revision: dict[str, Any]) -> str:
    kind_labels = {
        "insertion": "Insertion",
        "deletion": "Deletion",
        "moveFrom": "Move from",
        "moveTo": "Move to",
        "formatChange": "Format change",
    }
    parts = [kind_labels.get(revision.get("kind"), "Revision")]
    if revision.get("author"):
        parts.append(f"Author: {revision['author']}")
    if revision.get("date"):
        parts.append(f"Date: {revision['date']}")
    return " | ".join(parts)


def paragraph_tag_name(style_id: str | None, style_name: str | None) -> str:
    key = f"{style_id or ''} {style_name or ''}".lower()

    if "title" in key or "titulo" in key:
        return "h1"
    if "subtitle" in key or "subtitulo" in key:
        return "h2"

    for level in range(1, 7):
        if f"heading {level}" in key or f"titulo {level}" in key:
            return f"h{level}"

    return "p"


def paragraph_class_name(style_id: str | None, style_name: str | None) -> str:
    parts = ["paragraph-block"]
    if style_id:
        parts.append(f"style-id-{slugify(style_id)}")
    if style_name:
        parts.append(f"style-name-{slugify(style_name)}")
    return " ".join(parts)


def paragraph_css(props: dict[str, Any]) -> dict[str, str]:
    css: dict[str, str] = {}
    align_map = {"both": "justify"}
    if props.get("align"):
        css["text-align"] = align_map.get(props["align"], props["align"])
    if props.get("marginTopPt") is not None:
        css["margin-top"] = pt(props["marginTopPt"])
    if props.get("marginBottomPt") is not None:
        css["margin-bottom"] = pt(props["marginBottomPt"])
    if props.get("marginLeftPt") is not None:
        css["margin-left"] = pt(props["marginLeftPt"])
    if props.get("marginRightPt") is not None:
        css["margin-right"] = pt(props["marginRightPt"])
    if props.get("textIndentPt") is not None:
        css["text-indent"] = pt(props["textIndentPt"])
    if props.get("lineHeightPt") is not None:
        css["line-height"] = pt(props["lineHeightPt"])
    if props.get("backgroundColor"):
        css["background-color"] = props["backgroundColor"]
    if props.get("borderTopSpacePt") is not None:
        css["padding-top"] = pt(props["borderTopSpacePt"])
    if props.get("borderBottomSpacePt") is not None:
        css["padding-bottom"] = pt(props["borderBottomSpacePt"])
    if props.get("borderLeftSpacePt") is not None:
        css["padding-left"] = pt(props["borderLeftSpacePt"])
    if props.get("borderRightSpacePt") is not None:
        css["padding-right"] = pt(props["borderRightSpacePt"])
    if props.get("borderCss"):
        css["__raw__"] = props["borderCss"]
    return css


def run_css(props: dict[str, Any]) -> dict[str, str]:
    css: dict[str, str] = {}
    if props.get("fontFamily"):
        css["font-family"] = props["fontFamily"]
    if props.get("fontSizePt") is not None:
        css["font-size"] = pt(props["fontSizePt"])
    if props.get("bold"):
        css["font-weight"] = "700"
    if props.get("italic"):
        css["font-style"] = "italic"
    if props.get("underline"):
        css["text-decoration-line"] = "underline"
    if props.get("color"):
        css["color"] = props["color"]
    if props.get("highlight"):
        css["background-color"] = props["highlight"]
    if props.get("backgroundColor"):
        css["background-color"] = props["backgroundColor"]
    return css


def cell_css(props: dict[str, Any], table_props: dict[str, Any]) -> dict[str, str]:
    css: dict[str, str] = {}
    if props.get("backgroundColor"):
        css["background-color"] = props["backgroundColor"]
    if props.get("widthPt") is not None:
        css["width"] = pt(props["widthPt"])
    for source, key in [
        ("cellPaddingTopPt", "padding-top"),
        ("cellPaddingRightPt", "padding-right"),
        ("cellPaddingBottomPt", "padding-bottom"),
        ("cellPaddingLeftPt", "padding-left"),
    ]:
        if table_props.get(source) is not None:
            css[key] = pt(table_props[source])
    if props.get("borderCss"):
        css["border"] = "none"
        css["__raw__"] = props["borderCss"]
    return css


def table_css(props: dict[str, Any]) -> dict[str, str]:
    css: dict[str, str] = {}
    if props.get("widthPt") is not None:
        css["width"] = pt(props["widthPt"])
        css["max-width"] = "100%"
    elif props.get("gridColsPt"):
        css["width"] = pt(sum(props["gridColsPt"]))
        css["max-width"] = "100%"
    if props.get("marginLeftPt") is not None:
        css["margin-left"] = pt(props["marginLeftPt"])
    if props.get("backgroundColor"):
        css["background-color"] = props["backgroundColor"]
    if props.get("borderCss"):
        outer, inside = split_table_border_css(props["borderCss"])
        if outer:
            css["__raw__"] = outer
        if inside:
            css["__inside__"] = inside
    return css


def table_row_css(props: dict[str, Any]) -> dict[str, str]:
    css: dict[str, str] = {}
    if props.get("heightPt") is not None:
        rule = props.get("heightRule")
        if rule == "exact":
            css["height"] = pt(props["heightPt"])
        elif rule == "atLeast":
            css["min-height"] = pt(props["heightPt"])
    return css


def render_text(value: str) -> str:
    return escape(value).replace("\n", "<br/>").replace("\t", "&emsp;")


def css_from_map(css: dict[str, str]) -> str:
    raw = css.pop("__raw__", None) if "__raw__" in css else None
    inside = css.pop("__inside__", None) if "__inside__" in css else None
    parts = [f"{key}: {value}" for key, value in css.items()]
    if raw:
        parts.append(raw)
    if inside:
        parts.append(inside)
    return "; ".join(parts)


def split_table_border_css(value: str) -> tuple[str, str]:
    outer_parts: list[str] = []
    inside_parts: list[str] = []

    for part in value.split(";"):
        chunk = part.strip()
        if not chunk:
            continue
        if chunk.startswith("inside-horizontal:"):
            inside_parts.append(f"--docx-inside-horizontal: {chunk.split(':', 1)[1].strip()}")
        elif chunk.startswith("inside-vertical:"):
            inside_parts.append(f"--docx-inside-vertical: {chunk.split(':', 1)[1].strip()}")
        else:
            outer_parts.append(f"border-{chunk}")

    return "; ".join(outer_parts), "; ".join(inside_parts)


def merge_dicts(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged = dict(base)
    for key, value in override.items():
        if value is not None:
            merged[key] = value
    return merged


def split_inline_nodes_on_first_tab(nodes: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]] | None:
    left: list[dict[str, Any]] = []
    right: list[dict[str, Any]] = []
    found_tab = False

    for node in nodes:
        if node["type"] != "text":
            if found_tab:
                right.append(node)
            else:
                left.append(node)
            continue

        value = node["value"]
        if "\t" not in value:
            target = right if found_tab else left
            target.append(node)
            continue

        before, after = value.split("\t", 1)
        if before:
            left.append({**node, "value": before})
        if after:
            right.append({**node, "value": after})
        found_tab = True

    if not found_tab:
        return None

    return {"left": left, "right": right}


def resolve_fill(node: ET.Element | None) -> str | None:
    if node is None:
        return None
    fill = attr(node, "fill")
    if fill and fill != "auto":
        return normalize_color(fill)
    color = attr(node, "color")
    if color and color not in {"auto", "000000"}:
        return normalize_color(color)
    return None


def highlight_to_css(value: str) -> str | None:
    palette = {
        "yellow": "#fff59d",
        "green": "#bbf7d0",
        "cyan": "#a5f3fc",
        "magenta": "#f5d0fe",
        "blue": "#bfdbfe",
        "red": "#fecaca",
        "darkYellow": "#fde68a",
        "darkBlue": "#93c5fd",
        "darkCyan": "#67e8f9",
        "darkGreen": "#86efac",
        "darkMagenta": "#e879f9",
        "darkRed": "#fca5a5",
        "darkGray": "#d1d5db",
        "lightGray": "#e5e7eb",
        "black": "#111827",
    }
    return palette.get(value)


def normalize_color(value: str) -> str:
    cleaned = value.strip().lstrip("#")
    if cleaned.lower() == "auto":
        return "#000000"
    if len(cleaned) == 6:
        return f"#{cleaned}"
    return value


def half_points_to_pt(value: str | None) -> float | None:
    if not value or not value.isdigit():
        return None
    return int(value) / 2


def twips_to_pt(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return round(int(value) / 20, 2)
    except ValueError:
        return None


def pt(value: float) -> str:
    text = f"{value:.2f}".rstrip("0").rstrip(".")
    return f"{text}pt"


def slugify(value: str) -> str:
    return "".join(char.lower() if char.isalnum() else "-" for char in value).strip("-")


def data_attr(name: str, value: str | None) -> str:
    return f'{name}="{escape(value)}"' if value else ""


def class_attr(value: str) -> str:
    return f'class="{escape(value)}"' if value else ""


def attr(node: ET.Element | None, name: str) -> str | None:
    if node is None:
        return None
    return node.attrib.get(qname(name))


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def qname(name: str) -> str:
    return f"{{{NS['w']}}}{name}"
