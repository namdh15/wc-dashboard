"""Convert huongdan.docx to huongdan.md, preserving images and all content."""

import re
import os
import shutil
import zipfile
from docx import Document
from docx.oxml.ns import qn

# Namespaces
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------------------------------------------------------------------
# Image extraction
# ---------------------------------------------------------------------------

def extract_images(docx_path, images_dir):
    """Extract all media files from docx into images_dir. Returns {target_ref: local_path}."""
    if os.path.exists(images_dir):
        shutil.rmtree(images_dir)
    os.makedirs(images_dir)

    mapping = {}
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if name.startswith("word/media/"):
                filename = os.path.basename(name)
                dest = os.path.join(images_dir, filename)
                with z.open(name) as src, open(dest, "wb") as dst:
                    dst.write(src.read())
                # target_ref stored in rels is like "media/image1.png"
                mapping[f"media/{filename}"] = dest
    return mapping


# ---------------------------------------------------------------------------
# Drawing → image path
# ---------------------------------------------------------------------------

def get_drawing_image_path(drawing_elem, doc_part, image_map):
    """Return local image path for a w:drawing element, or None."""
    # Find the a:blip inside the drawing
    blip = drawing_elem.find(f".//{{{NS_A}}}blip")
    if blip is None:
        return None
    r_embed = blip.get(f"{{{NS_R}}}embed")
    if not r_embed:
        return None
    rel = doc_part.rels.get(r_embed)
    if rel is None:
        return None
    return image_map.get(rel.target_ref)


def get_drawing_description(drawing_elem):
    """Try to get alt-text / description from the drawing."""
    # docPr element has name/descr attributes
    docPr = drawing_elem.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr")
    if docPr is not None:
        descr = docPr.get("descr") or docPr.get("name") or ""
        return descr.strip()
    return ""


# ---------------------------------------------------------------------------
# Inline run text
# ---------------------------------------------------------------------------

def get_run_text(run):
    text = run.text
    if not text:
        return ""
    bold = run.bold
    italic = run.italic
    if bold and italic:
        return f"***{text}***"
    if bold:
        return f"**{text}**"
    if italic:
        return f"*{text}*"
    return text


# ---------------------------------------------------------------------------
# Numbering helpers
# ---------------------------------------------------------------------------

def _get_num_fmt(para, num_id, ilvl):
    """Return the numFmt value ('bullet', 'decimal', etc.) for this list level."""
    try:
        numbering_part = para.part.numbering_part
        if numbering_part is None:
            return ""
        numbering = numbering_part._element
        for num in numbering.findall(qn("w:num")):
            if num.get(qn("w:numId")) == num_id:
                abstract_ref = num.find(qn("w:abstractNumId"))
                if abstract_ref is None:
                    return ""
                abstract_id = abstract_ref.get(qn("w:val"))
                for abs_num in numbering.findall(qn("w:abstractNum")):
                    if abs_num.get(qn("w:abstractNumId")) == abstract_id:
                        for lvl in abs_num.findall(qn("w:lvl")):
                            if lvl.get(qn("w:ilvl")) == str(ilvl):
                                fmt_elem = lvl.find(qn("w:numFmt"))
                                if fmt_elem is not None:
                                    return fmt_elem.get(qn("w:val"), "")
        return ""
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Paragraph → markdown lines (may return multiple lines for mixed text+image)
# ---------------------------------------------------------------------------

def para_to_lines(para, list_counters, doc_part, image_map, images_dir_rel):
    """Convert a paragraph to a list of markdown lines."""
    style = para.style.name if para.style else ""
    block = para._element

    # Collect parts in document order: text runs and drawings
    parts = []
    for child in block.iter():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "r":
            # It's a run — check if it has text
            t = child.find(qn("w:t"))
            if t is not None and t.text:
                from docx.text.run import Run
                run = Run(child, para)
                parts.append(("text", get_run_text(run)))
        elif tag == "drawing":
            img_path = get_drawing_image_path(child, doc_part, image_map)
            if img_path:
                descr = get_drawing_description(child)
                filename = os.path.basename(img_path)
                rel_path = os.path.join(images_dir_rel, filename).replace("\\", "/")
                parts.append(("image", rel_path, descr))

    # Separate text parts from image parts
    text_parts = [p for p in parts if p[0] == "text"]
    image_parts = [p for p in parts if p[0] == "image"]
    inline = "".join(p[1] for p in text_parts)

    # Determine heading level
    heading_map = {
        "Heading 1": "# ", "Heading 2": "## ", "Heading 3": "### ",
        "Heading 4": "#### ", "Heading 5": "##### ", "Heading 6": "###### ",
        "Title": "# ", "Subtitle": "## ",
    }
    heading_prefix = None
    for style_key, prefix in heading_map.items():
        if style.startswith(style_key):
            list_counters.clear()
            heading_prefix = prefix
            break

    # Determine list prefix
    list_prefix = None
    num_pr = block.find(qn("w:pPr"))
    num_id = None
    ilvl = 0
    if num_pr is not None:
        num_elem = num_pr.find(qn("w:numPr"))
        if num_elem is not None:
            ilvl_elem = num_elem.find(qn("w:ilvl"))
            num_id_elem = num_elem.find(qn("w:numId"))
            if ilvl_elem is not None:
                ilvl = int(ilvl_elem.get(qn("w:val"), 0))
            if num_id_elem is not None:
                num_id = num_id_elem.get(qn("w:val"))

    if num_id is not None:
        indent = "  " * ilvl
        num_fmt = _get_num_fmt(para, num_id, ilvl)
        if num_fmt == "bullet":
            list_prefix = f"{indent}- "
        else:
            key = (num_id, ilvl)
            list_counters[key] = list_counters.get(key, 0) + 1
            list_prefix = f"{indent}{list_counters[key]}. "
    elif "List Bullet" in style:
        ilvl_n = int(re.search(r"\d", style).group()) - 1 if re.search(r"\d", style) else 0
        list_prefix = "  " * ilvl_n + "- "
    elif "List Number" in style:
        ilvl_n = int(re.search(r"\d", style).group()) - 1 if re.search(r"\d", style) else 0
        key = ("style", style)
        list_counters[key] = list_counters.get(key, 0) + 1
        list_prefix = "  " * ilvl_n + f"{list_counters[key]}. "

    # Build output lines
    lines = []

    # Text line
    text_stripped = inline.strip()
    if heading_prefix:
        if text_stripped:
            lines.append(heading_prefix + text_stripped)
    elif list_prefix:
        if text_stripped:
            lines.append(list_prefix + text_stripped)
        # Image inside a list item goes on next line indented
        for _, rel_path, descr in image_parts:
            lines.append(f"  ![{descr}]({rel_path})")
        return lines  # already handled images
    else:
        if text_stripped:
            lines.append(text_stripped)

    # Standalone images (paragraph-level)
    for _, rel_path, descr in image_parts:
        lines.append(f"![{descr}]({rel_path})")

    return lines


# ---------------------------------------------------------------------------
# Main conversion
# ---------------------------------------------------------------------------

def docx_to_md(docx_path, md_path, images_dir="images"):
    doc = Document(docx_path)
    image_map = extract_images(docx_path, images_dir)
    print(f"Extracted {len(image_map)} images to '{images_dir}/'")

    list_counters = {}
    all_lines = []

    for block in doc.element.body:
        tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag

        if tag == "p":
            from docx.text.paragraph import Paragraph
            para = Paragraph(block, doc)
            lines = para_to_lines(para, list_counters, doc.part, image_map, images_dir)
            if lines:
                all_lines.extend(lines)
            else:
                all_lines.append("")  # blank line to preserve spacing

        elif tag == "tbl":
            from docx.table import Table
            tbl = Table(block, doc)
            list_counters.clear()
            if all_lines and all_lines[-1] != "":
                all_lines.append("")
            all_lines.extend(table_to_md_lines(tbl))
            all_lines.append("")

    # Collapse 3+ consecutive blank lines → 1
    md = "\n".join(all_lines)
    md = re.sub(r"\n{3,}", "\n\n", md).strip()

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md)
    print(f"Saved: {md_path}")


def table_to_md_lines(table):
    rows = table.rows
    if not rows:
        return []
    lines = []
    for i, row in enumerate(rows):
        cells = []
        prev = None
        for cell in row.cells:
            text = " ".join(p.text.strip() for p in cell.paragraphs)
            if text != prev:
                cells.append(text)
                prev = text
            else:
                cells.append("")
        lines.append("| " + " | ".join(cells) + " |")
        if i == 0:
            lines.append("| " + " | ".join(["---"] * len(cells)) + " |")
    return lines


if __name__ == "__main__":
    docx_to_md("huongdan.docx", "huongdan.md", images_dir="images")
