"""
md_to_docx - Markdown to styled .docx converter.

Reads a Markdown source, converts it to OOXML using the paragraph and table
styles of a host .docx, and injects the result between two marker paragraphs
of that host: <!-- INSERT-START --> and <!-- INSERT-END -->. The host carries
everything that is not content: cover page, table of contents, legal
sections, headers and footers, and the styles themselves.

Supported Markdown subset:
    - ATX headings (# to ######) -> heading styles (numbered by the host)
    - Front-matter headings: `# Title {.frontmatter}` -> non-numbered styles
    - A page break is inserted before every numbered level-1 heading, except
      the first one.
    - Paragraphs, bullet lists (`- item`), ordered lists (`1. item`)
    - GFM pipe tables (header row bold)
    - Instruction callouts: `> [!INSTRUCTION] text` -> italic paragraph
    - Inline emphasis (*foo*, **foo**) within paragraphs and table cells
    - Images: ![alt](relative/path.png), embedded inline and scaled to the
      page content width if larger. Formats: PNG, JPEG. Optional manual
      width: ![alt](path.png){width=60%}.
    - Captions: a paragraph made of a single italic span starting with
      "Figure:" or "Table:" becomes a caption paragraph with a SEQ field, so
      the lists of figures / tables of the host fill in when fields are
      updated. An image followed by its caption stays on the same page.
    - "[Auto]" alone in a table cell or paragraph -> NUMPAGES field.
    - Layout markers, written as HTML comments on a line of their own (they
      are invisible in the rendered Markdown):
        <!-- pagebreak -->            page break
        <!-- linebreak -->            empty line
        <!-- colwidths: 20 50 30 -->  relative column widths of the table
                                      that directly follows
        <!-- internal:start --> ... <!-- internal:end -->
                                      the lines in between are left out
    - Code: when the host defines a paragraph style named "Code", fenced code
      blocks use it (one paragraph per line); when it defines the character
      style "Code Char" or "Code", inline code uses it. Otherwise code is
      plain text.

The style ids used for each construct are those of a standard Word document.
A host built on another template declares its own ids in a small JSON file
passed with --styles (see DEFAULT_STYLES below for the keys).

Out of scope:
    - Cell merging in tables
    - Nested lists or tables inside table cells
    - Footnotes
    - HTML passthrough (other than <br> and the markers above)
    - Cross-references and field codes other than caption SEQ fields

Usage:
    python md_to_docx.py --md source.md --host-docx host.docx --out result.docx
                         [--styles styles.json] [--keep-internal]
"""

from __future__ import annotations

import argparse
import json
import re
import struct
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from html import escape
from pathlib import Path
from typing import Iterable, Optional

from markdown_it import MarkdownIt
from markdown_it.token import Token

__version__ = "1.0.0"


# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

# Style ids and page geometry. The defaults are those of a standard Word
# document; a host built on another template overrides them with --styles.
DEFAULT_STYLES = {
    # Numbered headings, by Markdown level.
    "heading": {"1": "Heading1", "2": "Heading2", "3": "Heading3",
                "4": "Heading4", "5": "Heading5", "6": "Heading6"},
    # Headings marked {.frontmatter}: not numbered, no page break.
    "frontmatter_heading": {"1": "Title", "2": "Subtitle", "3": "Subtitle", "4": "Subtitle"},
    "body": "BodyText",
    "bullet": "ListBullet",
    "number": "ListNumber",
    "caption": "Caption",
    "table": "TableGrid",
    # Name (not id) of the host style used for code, see host_code_styles().
    "code_style_name": "Code",
    # Width of the text area, in twips (1/20 of a point). Default: an A4 page
    # (11906 twips wide) with left and right margins of about 2 cm.
    "content_width_twips": 9620,
}

NUMBERED_STYLE: dict = {}
FRONTMATTER_STYLE: dict = {}
BODY_STYLE = BULLET_STYLE = NUMBER_STYLE = CAPTION_STYLE = TABLE_STYLE = ""
CODE_STYLE_NAME = "Code"
TABLE_WIDTH_TWIPS = 0
MAX_IMAGE_WIDTH_EMU = 0


def apply_styles(overrides=None) -> None:
    """Set the style ids and page geometry used by the conversion: the
    defaults, updated by the given overrides (same keys as DEFAULT_STYLES)."""
    global NUMBERED_STYLE, FRONTMATTER_STYLE, BODY_STYLE, BULLET_STYLE, NUMBER_STYLE
    global CAPTION_STYLE, TABLE_STYLE, CODE_STYLE_NAME, TABLE_WIDTH_TWIPS, MAX_IMAGE_WIDTH_EMU
    styles = {key: (dict(value) if isinstance(value, dict) else value)
              for key, value in DEFAULT_STYLES.items()}
    for key, value in (overrides or {}).items():
        if key.startswith("_"):
            continue  # free-form comment keys
        if key not in styles:
            raise ValueError(f"unknown style key: {key}")
        if isinstance(styles[key], dict):
            styles[key].update({str(k): v for k, v in value.items()})
        else:
            styles[key] = value
    NUMBERED_STYLE = {int(k): v for k, v in styles["heading"].items()}
    FRONTMATTER_STYLE = {int(k): v for k, v in styles["frontmatter_heading"].items()}
    BODY_STYLE, BULLET_STYLE, NUMBER_STYLE = styles["body"], styles["bullet"], styles["number"]
    CAPTION_STYLE, TABLE_STYLE = styles["caption"], styles["table"]
    CODE_STYLE_NAME = styles["code_style_name"]
    TABLE_WIDTH_TWIPS = int(styles["content_width_twips"])
    # 1 twip = 635 EMU. An image is never wider than the text area.
    MAX_IMAGE_WIDTH_EMU = TABLE_WIDTH_TWIPS * 635


apply_styles()

# Default DPI assumed when image has no explicit physical resolution.
DEFAULT_DPI = 96
EMU_PER_INCH = 914400

START_MARKER = "<!-- INSERT-START -->"
END_MARKER = "<!-- INSERT-END -->"

# Relationship ID space we use for inserted images. Starting high
# guarantees no collision with the host's existing rIds (1..50 range
# in practice).
IMAGE_REL_ID_BASE = 1000


# --------------------------------------------------------------------------- #
# Image handling
# --------------------------------------------------------------------------- #


@dataclass
class ImageRef:
    """An image referenced from the Markdown that needs to be embedded."""
    source_path: Path
    media_filename: str  # e.g. "image_inj_001.png"
    rel_id: str          # e.g. "rId1001"
    width_emu: int
    height_emu: int


def _read_image_size(path: Path) -> tuple[int, int]:
    """Return image (width_px, height_px). Supports PNG and JPEG."""
    data = path.read_bytes()
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        # PNG: IHDR chunk starts at byte 8, length(4) + type(4) + width(4) + height(4)
        w = struct.unpack(">I", data[16:20])[0]
        h = struct.unpack(">I", data[20:24])[0]
        return w, h
    if data[:2] == b"\xff\xd8":
        # JPEG: walk markers to find SOF0/SOF2
        i = 2
        n = len(data)
        while i < n:
            if data[i] != 0xFF:
                raise ValueError(f"JPEG parse error in {path}")
            marker = data[i + 1]
            i += 2
            if marker in (0xD8, 0xD9):
                continue
            if 0xD0 <= marker <= 0xD7:
                continue
            seg_len = struct.unpack(">H", data[i:i + 2])[0]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3,
                          0xC5, 0xC6, 0xC7,
                          0xC9, 0xCA, 0xCB,
                          0xCD, 0xCE, 0xCF):
                # SOF marker: precision(1), height(2), width(2)
                h = struct.unpack(">H", data[i + 3:i + 5])[0]
                w = struct.unpack(">H", data[i + 5:i + 7])[0]
                return w, h
            i += seg_len
        raise ValueError(f"JPEG SOF marker not found in {path}")
    # Fallback: try Pillow if installed.
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size
    except ImportError as exc:
        raise ValueError(
            f"Unsupported image format for {path}; install Pillow for broader support."
        ) from exc


def _image_extension(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in (".jpg", ".jpeg"):
        return "jpeg"
    if suffix == ".png":
        return "png"
    return suffix.lstrip(".") or "bin"


def _compute_image_emu(
    width_px: int, height_px: int, width_pct: Optional[int] = None
) -> tuple[int, int]:
    """Convert pixel size to EMU. Without width_pct: cap at MAX_IMAGE_WIDTH_EMU.
    With width_pct (1..100): scale to that percentage of the page content width."""
    width_emu = int(width_px / DEFAULT_DPI * EMU_PER_INCH)
    height_emu = int(height_px / DEFAULT_DPI * EMU_PER_INCH)
    if width_pct:
        target = int(MAX_IMAGE_WIDTH_EMU * min(max(width_pct, 1), 100) / 100)
        scale = target / width_emu
        return target, int(height_emu * scale)
    if width_emu > MAX_IMAGE_WIDTH_EMU:
        scale = MAX_IMAGE_WIDTH_EMU / width_emu
        width_emu = MAX_IMAGE_WIDTH_EMU
        height_emu = int(height_emu * scale)
    return width_emu, height_emu


def _drawing_xml(image: ImageRef, doc_pr_id: int) -> str:
    """Build the <w:drawing>...</w:drawing> OOXML for an inline image."""
    return (
        '<w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<wp:extent cx="{image.width_emu}" cy="{image.height_emu}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_pr_id}" name="Picture {doc_pr_id}"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic>'
        '<pic:nvPicPr>'
        f'<pic:cNvPr id="{doc_pr_id}" name="{escape(image.media_filename)}"/>'
        '<pic:cNvPicPr/>'
        '</pic:nvPicPr>'
        '<pic:blipFill>'
        f'<a:blip r:embed="{image.rel_id}"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</pic:blipFill>'
        '<pic:spPr>'
        '<a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{image.width_emu}" cy="{image.height_emu}"/>'
        '</a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr>'
        '</pic:pic>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing>'
    )


# --------------------------------------------------------------------------- #
# Conversion state
# --------------------------------------------------------------------------- #


@dataclass
class ConvertState:
    """State threaded through the conversion to track images and headings."""
    md_dir: Path
    images: list[ImageRef] = field(default_factory=list)
    next_doc_pr_id: int = 1
    next_rel_id: int = IMAGE_REL_ID_BASE
    has_emitted_content: bool = False
    after_heading: bool = False
    after_manual_break: bool = False
    # Column widths announced by a <!-- colwidths: ... --> marker, for the next table.
    pending_col_widths: Optional[list[float]] = None
    # Style ids of the host's "Code" styles, when the host defines them.
    code_para_style: Optional[str] = None
    code_char_style: Optional[str] = None
    # Running SEQ numbers per caption kind (placeholder values; Word
    # recomputes them when fields are updated).
    seq: dict = field(default_factory=lambda: {"Figure": 0, "Table": 0})

    def register_image(self, src_relpath: str, width_pct: Optional[int] = None) -> ImageRef:
        path = (self.md_dir / src_relpath).resolve()
        if not path.exists():
            raise FileNotFoundError(f"Image not found: {path}")
        w_px, h_px = _read_image_size(path)
        w_emu, h_emu = _compute_image_emu(w_px, h_px, width_pct)
        seq = len(self.images) + 1
        media_filename = f"image_inj_{seq:03d}.{_image_extension(path)}"
        rel_id = f"rId{self.next_rel_id}"
        self.next_rel_id += 1
        ref = ImageRef(
            source_path=path,
            media_filename=media_filename,
            rel_id=rel_id,
            width_emu=w_emu,
            height_emu=h_emu,
        )
        self.images.append(ref)
        return ref


# --------------------------------------------------------------------------- #
# OOXML builders
# --------------------------------------------------------------------------- #


def _xml_text(text: str) -> str:
    return escape(text, quote=False)


def _run(text: str, *, italic: bool = False, bold: bool = False,
         char_style: Optional[str] = None) -> str:
    rpr_parts = []
    if char_style:
        # Schema order in w:rPr: rStyle comes first.
        rpr_parts.append(f'<w:rStyle w:val="{char_style}"/>')
    if bold:
        rpr_parts.append("<w:b/>")
    if italic:
        rpr_parts.append("<w:i/>")
    rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>" if rpr_parts else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{_xml_text(text)}</w:t></w:r>'


def _paragraph(
    style: str,
    content_xml: str = "",
    *,
    keep_next: bool = False,
    keep_lines: bool = False,
) -> str:
    # Schema order in w:pPr: pStyle, keepNext, keepLines.
    props = [f'<w:pStyle w:val="{style}"/>']
    if keep_next:
        props.append("<w:keepNext/>")
    if keep_lines:
        props.append("<w:keepLines/>")
    return f'<w:p><w:pPr>{"".join(props)}</w:pPr>{content_xml}</w:p>'


def _heading_paragraph(style_id: str, content_xml: str) -> str:
    # keepNext: a heading never sits alone at the bottom of a page — it moves
    # with whatever follows it. keepLines: a multi-line heading never splits.
    return _paragraph(style_id, content_xml, keep_next=True, keep_lines=True)


def _page_break_paragraph() -> str:
    """Empty paragraph carrying a page break."""
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


# --------------------------------------------------------------------------- #
# Inline run rendering
# --------------------------------------------------------------------------- #


def _runs_from_inline(
    tokens: list[Token],
    state: ConvertState,
    *,
    force_italic: bool = False,
    force_bold: bool = False,
) -> str:
    """Turn an inline token stream into <w:r> runs (and <w:drawing> for images)."""
    out: list[str] = []
    italic = force_italic
    bold = force_bold
    skip = -1  # index of a text token consumed as an image width attribute

    for idx, tok in enumerate(tokens):
        ttype = tok.type
        if idx == skip:
            continue
        if ttype == "text":
            if AUTO_PAGES_RE.match(tok.content) and len(tokens) == 1:
                out.append(_numpages_field())
            else:
                out.append(_run(tok.content, italic=italic, bold=bold))
        elif ttype == "em_open":
            if not force_italic:
                italic = True
        elif ttype == "em_close":
            if not force_italic:
                italic = False
        elif ttype == "strong_open":
            if not force_bold:
                bold = True
        elif ttype == "strong_close":
            if not force_bold:
                bold = False
        elif ttype == "softbreak":
            out.append(_run(" ", italic=italic, bold=bold))
        elif ttype == "hardbreak":
            out.append('<w:r><w:br/></w:r>')
        elif ttype == "code_inline":
            out.append(_run(tok.content, italic=italic, bold=bold,
                            char_style=state.code_char_style))
        elif ttype == "html_inline":
            # <br> becomes a line break (notably inside table cells); comments
            # and any other raw HTML never reach the docx.
            if re.fullmatch(r"<br\s*/?\s*>", tok.content.strip(), re.IGNORECASE):
                out.append("<w:r><w:br/></w:r>")
        elif ttype == "image":
            src = tok.attrs.get("src", "") if tok.attrs else ""
            if not src:
                continue
            width_pct = None
            if idx + 1 < len(tokens) and tokens[idx + 1].type == "text":
                m = IMAGE_WIDTH_RE.match(tokens[idx + 1].content)
                if m:
                    width_pct = int(m.group(1))
                    rest = tokens[idx + 1].content[m.end():]
                    if rest.strip():
                        tokens[idx + 1].content = rest
                    else:
                        skip = idx + 1
            image = state.register_image(str(src), width_pct)
            doc_pr_id = state.next_doc_pr_id
            state.next_doc_pr_id += 1
            out.append(f'<w:r>{_drawing_xml(image, doc_pr_id)}</w:r>')
        else:
            if tok.content:
                out.append(_run(tok.content, italic=italic, bold=bold))

    return "".join(out)


# --------------------------------------------------------------------------- #
# Heading handling — frontmatter detection
# --------------------------------------------------------------------------- #


FRONTMATTER_RE = re.compile(r"\s*\{\.frontmatter\}\s*$")
IMAGE_WIDTH_RE = re.compile(r"\s*\{width=(\d{1,3})%\}")
CAPTION_RE = re.compile(r"^(Figure|Table)\s*:\s*", re.IGNORECASE)
# "[Auto]" on its own (typically the Page count cell of the Document control
# table) becomes a NUMPAGES field.
AUTO_PAGES_RE = re.compile(r"^\s*\[auto\]\s*$", re.IGNORECASE)


def _numpages_field() -> str:
    return ('<w:fldSimple w:instr=" NUMPAGES ">'
            '<w:r><w:rPr><w:noProof/></w:rPr><w:t>0</w:t></w:r></w:fldSimple>')


# --------------------------------------------------------------------------- #
# Captions
# --------------------------------------------------------------------------- #


def _caption_of(children: list[Token]) -> Optional[tuple[str, str]]:
    """(kind, text) when the inline stream is exactly one italic span whose
    text starts with "Figure:" or "Table:"; None otherwise."""
    if len(children) < 3 or children[0].type != "em_open" or children[-1].type != "em_close":
        return None
    parts = []
    for tok in children[1:-1]:
        if tok.type == "text" or tok.type == "code_inline":
            parts.append(tok.content)
        elif tok.type == "softbreak":
            parts.append(" ")
        else:
            return None
    text = "".join(parts).strip()
    m = CAPTION_RE.match(text)
    if not m:
        return None
    kind = m.group(1).capitalize()
    return kind, text[m.end():].strip()


def _caption_paragraph(kind: str, text: str, state: ConvertState, *, keep_next: bool = False) -> str:
    state.seq[kind] += 1
    runs = (
        _run(f"{kind} ")
        + f'<w:fldSimple w:instr=" SEQ {kind} \\* ARABIC ">'
        f'<w:r><w:rPr><w:noProof/></w:rPr><w:t>{state.seq[kind]}</w:t></w:r></w:fldSimple>'
        + _run(f": {text}")
    )
    return _paragraph(CAPTION_STYLE, runs, keep_next=keep_next)


def _is_caption_paragraph(tokens: list[Token], i: int) -> bool:
    """True when tokens[i] opens a paragraph that is a caption."""
    return (i + 1 < len(tokens) and tokens[i].type == "paragraph_open"
            and _caption_of(list(tokens[i + 1].children or [])) is not None)


def _heading_is_frontmatter_and_strip(children: list[Token]) -> bool:
    """
    If the heading's last text token ends with `{.frontmatter}`, strip the
    marker (mutating the token's content) and return True.
    """
    # Walk from the end to find the last text-bearing token.
    for tok in reversed(children):
        if tok.type == "text":
            m = FRONTMATTER_RE.search(tok.content)
            if m:
                tok.content = tok.content[: m.start()]
                # If the resulting content is empty, the caller will simply
                # render no run for this token — which is fine.
                return True
            return False
        if tok.type in ("strong_close", "em_close", "code_inline"):
            # Skip closing tokens; keep looking.
            continue
        # Any other token type without text in tail → not a frontmatter heading.
        break
    return False


# --------------------------------------------------------------------------- #
# Block-level walker
# --------------------------------------------------------------------------- #


MARKER = re.compile(r"<!--\s*(page[- ]?break|line[- ]?break)\s*-->", re.IGNORECASE)
COLWIDTHS = re.compile(r"<!--\s*col[- ]?widths\s*:\s*([\d.,;%\s]+?)\s*-->", re.IGNORECASE)


def _walk(tokens: list[Token], state: ConvertState) -> Iterable[str]:
    i = 0
    n = len(tokens)
    while i < n:
        tok = tokens[i]
        ttype = tok.type

        # True only for the first block emitted right after a heading.
        after_heading = state.after_heading
        state.after_heading = False

        # True only for the block that directly follows a manual page break.
        after_manual_break = state.after_manual_break
        state.after_manual_break = False

        # Widths announced by a marker apply to the table that directly follows.
        col_widths = state.pending_col_widths
        state.pending_col_widths = None

        # --- Headings ------------------------------------------------------ #
        if ttype == "heading_open":
            level = int(tok.tag[1])
            inline = tokens[i + 1]
            children = list(inline.children or [])
            is_frontmatter = _heading_is_frontmatter_and_strip(children)
            runs = _runs_from_inline(children, state)
            if is_frontmatter:
                style = FRONTMATTER_STYLE.get(level, NUMBERED_STYLE.get(level, "Heading1"))
                yield _heading_paragraph(style, runs)
            else:
                # Page break before every numbered Heading1, EXCEPT when
                # nothing has been emitted yet (i.e. the document starts
                # directly with an H1 — the host's INSERT-START already
                # sits at a clean page boundary).
                if level == 1 and state.has_emitted_content and not after_manual_break:
                    yield _page_break_paragraph()
                style = NUMBERED_STYLE.get(level, "Heading1")
                yield _heading_paragraph(style, runs)
            state.has_emitted_content = True
            state.after_heading = True
            i += 3
            continue

        # --- Block quotes / instruction callouts -------------------------- #
        if ttype == "blockquote_open":
            j = i + 1
            depth = 1
            while j < n:
                if tokens[j].type == "blockquote_open":
                    depth += 1
                elif tokens[j].type == "blockquote_close":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            yield from _handle_blockquote(tokens[i + 1 : j], state)
            i = j + 1
            continue

        # --- Paragraph ---------------------------------------------------- #
        if ttype == "paragraph_open":
            inline = tokens[i + 1]
            children = list(inline.children or [])
            caption = _caption_of(children)
            if caption:
                # A caption placed before a table stays with it; a caption
                # placed after an image or table simply follows it.
                next_is_table = i + 3 < n and tokens[i + 3].type == "table_open"
                yield _caption_paragraph(caption[0], caption[1], state, keep_next=next_is_table)
                state.has_emitted_content = True
                i += 3
                continue
            runs = _runs_from_inline(children, state)
            has_image = any(t.type == "image" for t in children)
            caption_follows = has_image and _is_caption_paragraph(tokens, i + 3)
            # The first paragraph after a heading is unsplittable (keepLines)
            # and chained to the start of the following block (keepNext): a
            # section may only open at the bottom of a page if the heading,
            # this paragraph AND the first line of the next block fit there —
            # otherwise the whole group moves to the top of the next page.
            # The keepNext chain stops at section boundaries: when the next
            # block is a heading, chaining into it would link section after
            # section (headings carry keepNext too) and Word silently drops
            # keep constraints it cannot satisfy within one page.
            next_is_heading = i + 3 < n and tokens[i + 3].type == "heading_open"
            yield _paragraph(
                BODY_STYLE,
                runs,
                keep_next=(after_heading and not next_is_heading) or caption_follows,
                keep_lines=after_heading,
            )
            state.has_emitted_content = True
            i += 3
            continue

        # --- Bullet list -------------------------------------------------- #
        if ttype == "bullet_list_open":
            j = i + 1
            depth = 1
            while j < n:
                if tokens[j].type == "bullet_list_open":
                    depth += 1
                elif tokens[j].type == "bullet_list_close":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            yield from _handle_list(
                tokens[i + 1 : j], state, style=BULLET_STYLE, keep_first=after_heading
            )
            i = j + 1
            continue

        # --- Ordered list ------------------------------------------------- #
        if ttype == "ordered_list_open":
            j = i + 1
            depth = 1
            while j < n:
                if tokens[j].type == "ordered_list_open":
                    depth += 1
                elif tokens[j].type == "ordered_list_close":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            yield from _handle_list(
                tokens[i + 1 : j], state, style=NUMBER_STYLE, keep_first=after_heading
            )
            i = j + 1
            continue

        # --- Table -------------------------------------------------------- #
        if ttype == "table_open":
            j = i + 1
            while j < n and tokens[j].type != "table_close":
                j += 1
            yield _handle_table(tokens[i + 1 : j], state, col_widths)
            i = j + 1
            continue

        # --- Horizontal rule --------------------------------------------- #
        if ttype == "hr":
            # Treat as a soft separator: empty paragraph.
            yield _paragraph(BODY_STYLE)
            i += 1
            continue

        # --- Raw HTML block (comments, etc.) ------------------------------ #
        if ttype == "html_block":
            # HTML passthrough is out of scope; block-level comments
            # (<!-- TODO ... -->) and stray HTML are silently dropped. A
            # dropped block emits nothing, so the after-heading state must
            # survive for the next real block.
            # Two comments are layout markers, invisible in the rendered
            # Markdown: <!-- pagebreak --> and <!-- linebreak -->.
            widths_marker = COLWIDTHS.fullmatch(tok.content.strip())
            if widths_marker:
                values = [float(v) for v in re.findall(r"\d+(?:\.\d+)?", widths_marker.group(1))]
                if values and all(v > 0 for v in values):
                    state.pending_col_widths = values
            marker = MARKER.fullmatch(tok.content.strip())
            if marker:
                kind = marker.group(1).lower().replace("-", "").replace(" ", "")
                if kind == "pagebreak":
                    yield _page_break_paragraph()
                    state.after_manual_break = True
                else:
                    yield _paragraph(BODY_STYLE)
                state.has_emitted_content = True
            state.after_heading = after_heading
            i += 1
            continue

        # --- Fence / code block ------------------------------------------ #
        if ttype in ("fence", "code_block"):
            content = tok.content.rstrip("\n")
            if state.code_para_style:
                # One paragraph per line, in the host's "Code" style.
                for code_line in content.split("\n"):
                    yield _paragraph(state.code_para_style, _run(code_line) if code_line else "")
            else:
                yield _paragraph(BODY_STYLE, _run(content))
            state.has_emitted_content = True
            i += 1
            continue

        i += 1


def _handle_blockquote(inner: list[Token], state: ConvertState) -> Iterable[str]:
    """Render a blockquote as italic BodyText paragraphs (instructional)."""
    i = 0
    n = len(inner)
    instruction_consumed = False

    while i < n:
        tok = inner[i]
        if tok.type == "paragraph_open":
            inline = inner[i + 1]
            children = list(inline.children or [])

            if not instruction_consumed and children:
                first = children[0]
                if first.type == "text":
                    m = re.match(r"^\s*\[!INSTRUCTION\]\s*", first.content)
                    if m:
                        first.content = first.content[m.end():]
                        if not first.content:
                            children = children[1:]
                instruction_consumed = True

            runs = _runs_from_inline(children, state, force_italic=True)
            if runs.strip():
                yield _paragraph(BODY_STYLE, runs)
            i += 3
            continue

        if tok.type == "blockquote_open":
            j = i + 1
            depth = 1
            while j < n:
                if inner[j].type == "blockquote_open":
                    depth += 1
                elif inner[j].type == "blockquote_close":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            yield from _handle_blockquote(inner[i + 1 : j], state)
            i = j + 1
            continue

        i += 1


def _handle_list(
    inner: list[Token], state: ConvertState, *, style: str, keep_first: bool = False
) -> Iterable[str]:
    i = 0
    n = len(inner)
    first_item = keep_first
    while i < n:
        tok = inner[i]
        if tok.type == "list_item_open":
            j = i + 1
            depth = 1
            while j < n:
                if inner[j].type == "list_item_open":
                    depth += 1
                elif inner[j].type == "list_item_close":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            item_inner = inner[i + 1 : j]
            first = True
            k = 0
            m = len(item_inner)
            while k < m:
                if item_inner[k].type == "paragraph_open":
                    inline = item_inner[k + 1]
                    runs = _runs_from_inline(list(inline.children or []), state)
                    p_style = style if first else BODY_STYLE
                    yield _paragraph(
                        p_style, runs, keep_next=first_item, keep_lines=first_item
                    )
                    first = False
                    first_item = False
                    k += 3
                    continue
                k += 1
            i = j + 1
            continue
        i += 1


def _handle_table(inner: list[Token], state: ConvertState, widths: Optional[list[float]] = None) -> str:
    header_cells: list[str] = []
    body_rows: list[list[str]] = []

    i = 0
    n = len(inner)
    in_thead = False
    in_tbody = False
    current_row: list[str] | None = None

    while i < n:
        tok = inner[i]
        t = tok.type
        if t == "thead_open":
            in_thead = True
        elif t == "thead_close":
            in_thead = False
        elif t == "tbody_open":
            in_tbody = True
        elif t == "tbody_close":
            in_tbody = False
        elif t == "tr_open":
            current_row = []
        elif t == "tr_close":
            if in_thead and current_row is not None:
                header_cells = current_row
            elif in_tbody and current_row is not None:
                body_rows.append(current_row)
            current_row = None
        elif t in ("th_open", "td_open"):
            inline = inner[i + 1]
            children = list(inline.children or [])
            if t == "th_open":
                runs = _runs_from_inline(children, state, force_bold=True)
            else:
                runs = _runs_from_inline(children, state)
            if current_row is not None:
                current_row.append(runs)
            i += 3
            continue
        i += 1

    return _build_table(header_cells, body_rows, widths)


def _build_table(header_runs: list[str], body_rows: list[list[str]],
                 widths: Optional[list[float]] = None) -> str:
    n_cols = len(header_runs)
    if n_cols == 0:
        return ""
    if widths and len(widths) != n_cols:
        print(f"warning: colwidths gives {len(widths)} value(s) for a table of {n_cols} "
              "column(s); equal widths used", file=sys.stderr)
        widths = None
    if widths:
        # Relative values (percentages or any proportions), scaled to the page width.
        total = sum(widths)
        col_widths = [int(TABLE_WIDTH_TWIPS * w / total) for w in widths[:-1]]
        col_widths.append(TABLE_WIDTH_TWIPS - sum(col_widths))
    else:
        base = TABLE_WIDTH_TWIPS // n_cols
        col_widths = [base] * (n_cols - 1) + [TABLE_WIDTH_TWIPS - base * (n_cols - 1)]
    # Explicit widths are only honoured by Word with a fixed table layout.
    layout = (f'<w:tblW w:w="{TABLE_WIDTH_TWIPS}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
              if widths else '<w:tblW w:w="0" w:type="auto"/>')
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)

    def _cell(width: int, runs_xml: str) -> str:
        para = f'<w:p>{runs_xml}</w:p>' if runs_xml else '<w:p/>'
        return (
            f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
            f'{para}</w:tc>'
        )

    header_cells_xml = "".join(
        _cell(w, runs) for w, runs in zip(col_widths, header_runs)
    )
    header_row = (
        '<w:tr><w:trPr>'
        '<w:cnfStyle w:val="100000000000" w:firstRow="1"/>'
        '<w:tblHeader/>'
        '</w:trPr>'
        f'{header_cells_xml}</w:tr>'
    )

    body_xml = ""
    for row in body_rows:
        if len(row) < n_cols:
            row = row + [""] * (n_cols - len(row))
        elif len(row) > n_cols:
            row = row[:n_cols]
        cells = "".join(_cell(w, runs) for w, runs in zip(col_widths, row))
        body_xml += f"<w:tr>{cells}</w:tr>"

    return (
        '<w:tbl>'
        '<w:tblPr>'
        f'<w:tblStyle w:val="{TABLE_STYLE}"/>'
        f'{layout}'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
        'w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'{header_row}{body_xml}'
        '</w:tbl>'
    )


# --------------------------------------------------------------------------- #
# Top-level conversion
# --------------------------------------------------------------------------- #


def _make_parser() -> MarkdownIt:
    # html=True so raw HTML becomes html_inline/html_block tokens: <br> is
    # rendered as a line break, comments and other HTML are dropped. With
    # html=False they would surface as literal text in the docx.
    md = MarkdownIt("commonmark", {"html": True, "breaks": False})
    md.enable("table")
    return md


def host_code_styles(host_path: Path) -> tuple[Optional[str], Optional[str]]:
    """Ids of the host styles used for code: the paragraph style named "Code"
    (fenced blocks) and the character style named "Code Char" or "Code"
    (inline code). None when the host does not define them."""
    try:
        with zipfile.ZipFile(host_path, "r") as zf:
            styles_xml = zf.read("word/styles.xml").decode("utf-8")
    except (KeyError, zipfile.BadZipFile, OSError):
        return None, None
    para_id = char_id = None
    for m in re.finditer(r"<w:style\b([^>]*)>(.*?)</w:style>", styles_xml, re.S):
        attrs, body = m.group(1), m.group(2)
        name = re.search(r'<w:name w:val="([^"]*)"', body)
        kind = re.search(r'w:type="([^"]*)"', attrs)
        style_id = re.search(r'w:styleId="([^"]*)"', attrs)
        if not (name and kind and style_id):
            continue
        label = name.group(1).strip().lower()
        wanted = CODE_STYLE_NAME.strip().lower()
        if kind.group(1) == "paragraph" and label == wanted:
            para_id = style_id.group(1)
        elif kind.group(1) == "character" and label in (wanted + " char", wanted):
            char_id = style_id.group(1)
    return para_id, char_id


INTERNAL_START = re.compile(r"^\s*<!--\s*internal\s*:\s*start\s*-->\s*$", re.IGNORECASE)
INTERNAL_END = re.compile(r"^\s*<!--\s*internal\s*:\s*end\s*-->\s*$", re.IGNORECASE)


def strip_internal_blocks(md_text: str) -> str:
    """Remove everything between <!-- internal:start --> and <!-- internal:end -->
    (markers on a line of their own). Such blocks stay in the Markdown source
    and are left out of the generated document. A block that is never closed
    is dropped up to the end of the document, with a warning: leaving content
    out is the safe side."""
    out, skipping, in_code, opened_at = [], False, False, 0
    for number, line in enumerate(md_text.splitlines(), start=1):
        if line.startswith("```") and not skipping:
            in_code = not in_code
        if not in_code:
            if not skipping and INTERNAL_START.match(line):
                skipping, opened_at = True, number
                continue
            if skipping and INTERNAL_END.match(line):
                skipping = False
                continue
        if not skipping:
            out.append(line)
    if skipping:
        print(f"warning: internal block opened at line {opened_at} is never closed; "
              "everything after it is left out", file=sys.stderr)
    return "\n".join(out) + "\n"


def convert_markdown(md_text: str, md_dir: Path,
                     code_styles: tuple[Optional[str], Optional[str]] = (None, None),
                     keep_internal: bool = False,
                     ) -> tuple[str, list[ImageRef]]:
    parser = _make_parser()
    # With keep_internal the blocks stay; their markers are plain comments
    # and are dropped like any other comment.
    tokens = parser.parse(md_text if keep_internal else strip_internal_blocks(md_text))
    state = ConvertState(md_dir=md_dir, code_para_style=code_styles[0],
                         code_char_style=code_styles[1])
    parts = list(_walk(tokens, state))
    return "\n".join(parts), state.images


# --------------------------------------------------------------------------- #
# Host injection
# --------------------------------------------------------------------------- #


def _add_image_relationships(
    rels_path: Path, images: list[ImageRef]
) -> None:
    """Append <Relationship> entries for each image to document.xml.rels."""
    if not images:
        return
    txt = rels_path.read_text(encoding="utf-8")
    if "</Relationships>" not in txt:
        raise ValueError(f"Malformed rels file: {rels_path}")
    new_rels = "".join(
        f'<Relationship Id="{img.rel_id}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="media/{img.media_filename}"/>'
        for img in images
    )
    txt = txt.replace("</Relationships>", f"{new_rels}</Relationships>")
    rels_path.write_text(txt, encoding="utf-8")


def _ensure_image_content_types(content_types_path: Path, images: list[ImageRef]) -> None:
    """Add Default content-type entries for image extensions used."""
    if not images:
        return
    txt = content_types_path.read_text(encoding="utf-8")
    extensions_used = {
        _image_extension(img.source_path) for img in images
    }
    additions = []
    for ext in extensions_used:
        if ext == "jpeg":
            # Media parts are always written as .jpeg; a host Default for "jpg"
            # does not cover them, so "jpeg" must be declared in its own right.
            tag = '<Default Extension="jpeg" ContentType="image/jpeg"/>'
            if tag not in txt:
                additions.append(tag)
        elif ext == "png":
            tag = '<Default Extension="png" ContentType="image/png"/>'
            if tag not in txt:
                additions.append(tag)
    if additions:
        txt = txt.replace(
            "</Types>", "".join(additions) + "</Types>"
        )
        content_types_path.write_text(txt, encoding="utf-8")


# Elements that must follow w:updateFields in w:settings (schema order).
_SETTINGS_AFTER_UPDATE_FIELDS = (
    "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids",
    "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
    "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
    "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
    "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
)


def _set_update_fields(settings_xml: str) -> str:
    """Ask Word to refresh every field (TOC, lists of figures / tables, SEQ
    numbers) when the document is opened: Word prompts once, the user
    confirms, and the document is up to date. Removes any existing
    w:updateFields first, then inserts the element at its schema position."""
    settings_xml = re.sub(r'<w:updateFields\b[^>]*/>', "", settings_xml)
    element = '<w:updateFields w:val="true"/>'
    for name in _SETTINGS_AFTER_UPDATE_FIELDS:
        m = re.search(rf"<w:{name}\b", settings_xml)
        if m:
            return settings_xml[:m.start()] + element + settings_xml[m.start():]
    return settings_xml.replace("</w:settings>", element + "</w:settings>")


def inject_into_host(
    host_path: Path,
    content_xml: str,
    images: list[ImageRef],
    output_path: Path,
    *,
    update_fields: bool = True,
) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(host_path, "r") as zf:
            zf.extractall(tmp_path)

        document_xml_path = tmp_path / "word" / "document.xml"
        if not document_xml_path.exists():
            raise FileNotFoundError(
                f"Host .docx has no word/document.xml: {host_path}"
            )
        doc = document_xml_path.read_text(encoding="utf-8")

        # A marker is either a raw XML comment inside w:t (hand-made hosts)
        # or, once Word has re-saved the host, plain escaped text.
        def _find_marker(marker: str) -> int:
            idx = doc.find(marker)
            if idx < 0:
                idx = doc.find(escape(marker, quote=False))
            return idx

        start_text_idx = _find_marker(START_MARKER)
        if start_text_idx < 0:
            raise ValueError(
                f"START_MARKER {START_MARKER!r} not found in host document.xml. "
                "Add a paragraph containing exactly that text in your host .docx."
            )
        end_text_idx = _find_marker(END_MARKER)
        if end_text_idx < 0:
            raise ValueError(
                f"END_MARKER {END_MARKER!r} not found in host document.xml. "
                "Add a paragraph containing exactly that text in your host .docx."
            )

        # The marker paragraph may be written <w:p> or <w:p attr...>: take the
        # nearest opening of either form, never an earlier paragraph.
        zone_start = max(doc.rfind("<w:p ", 0, start_text_idx),
                         doc.rfind("<w:p>", 0, start_text_idx))
        if zone_start < 0:
            raise ValueError("Could not locate <w:p ...> opening for start marker.")

        zone_end_close = doc.find("</w:p>", end_text_idx)
        if zone_end_close < 0:
            raise ValueError("Could not locate </w:p> closing for end marker.")
        zone_end = zone_end_close + len("</w:p>")

        new_doc = doc[:zone_start] + content_xml + doc[zone_end:]
        document_xml_path.write_text(new_doc, encoding="utf-8")

        # Strip local-path attachedTemplate references.
        rels = tmp_path / "word" / "_rels" / "settings.xml.rels"
        if rels.exists():
            txt = rels.read_text(encoding="utf-8")
            txt = re.sub(
                r'<Relationship\b[^>]*\battachedTemplate\b[^>]*/>',
                "",
                txt,
            )
            rels.write_text(txt, encoding="utf-8")
        settings = tmp_path / "word" / "settings.xml"
        if settings.exists():
            txt = settings.read_text(encoding="utf-8")
            txt = re.sub(r'<w:attachedTemplate\b[^/]*/>', "", txt)
            if update_fields:
                txt = _set_update_fields(txt)
            settings.write_text(txt, encoding="utf-8")

        # Copy image binaries and register relationships + content-types.
        if images:
            media_dir = tmp_path / "word" / "media"
            media_dir.mkdir(exist_ok=True)
            for img in images:
                target = media_dir / img.media_filename
                target.write_bytes(img.source_path.read_bytes())
            _add_image_relationships(
                tmp_path / "word" / "_rels" / "document.xml.rels", images
            )
            _ensure_image_content_types(
                tmp_path / "[Content_Types].xml", images
            )

        if output_path.exists():
            output_path.unlink()
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in tmp_path.rglob("*"):
                if path.is_file():
                    arcname = path.relative_to(tmp_path).as_posix()
                    zf.write(path, arcname)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="md_to_docx",
        description=(
            f"md_to_docx v{__version__}. Convert a Markdown file to a styled .docx "
            "by injecting it into a host .docx that contains the INSERT-START "
            "and INSERT-END markers."
        ),
    )
    parser.add_argument("--md", required=True, type=Path,
                        help="Path to the source Markdown file.")
    parser.add_argument("--host-docx", required=True, type=Path,
                        help="Path to the host .docx (with the two markers).")
    parser.add_argument("--out", required=True, type=Path,
                        help="Path to the output .docx.")
    parser.add_argument("--no-update-fields", action="store_true",
                        help="Do not ask Word to update all fields (TOC, lists, "
                             "captions) when the document is opened.")
    parser.add_argument("--styles", type=Path,
                        help="JSON file overriding the style ids and the page "
                             "geometry (keys of DEFAULT_STYLES).")
    parser.add_argument("--keep-internal", action="store_true",
                        help="Keep the blocks enclosed by the internal:start / "
                             "internal:end markers (internal edition of a document).")
    parser.add_argument("--version", action="version",
                        version=f"%(prog)s {__version__}")
    args = parser.parse_args()

    if not args.md.exists():
        print(f"error: --md not found: {args.md}", file=sys.stderr)
        return 2
    if not args.host_docx.exists():
        print(f"error: --host-docx not found: {args.host_docx}", file=sys.stderr)
        return 2

    overrides = None
    if args.styles:
        if not args.styles.exists():
            print(f"error: --styles not found: {args.styles}", file=sys.stderr)
            return 2
        overrides = json.loads(args.styles.read_text(encoding="utf-8"))
    try:
        apply_styles(overrides)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    md_text = args.md.read_text(encoding="utf-8")
    code_styles = host_code_styles(args.host_docx)
    content_xml, images = convert_markdown(md_text, args.md.parent.resolve(), code_styles,
                                           keep_internal=args.keep_internal)
    inject_into_host(args.host_docx, content_xml, images, args.out,
                     update_fields=not args.no_update_fields)
    n_img = len(images)
    img_msg = f" ({n_img} image{'s' if n_img != 1 else ''})" if n_img else ""
    print(f"wrote {args.out}{img_msg}")
    if code_styles[0] or code_styles[1]:
        print(f"code styles: block={code_styles[0] or '-'}, inline={code_styles[1] or '-'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
