#!/usr/bin/env python3
"""Build a host template (with the INSERT-START / INSERT-END markers) from an
existing, fully written Word document.

The source document is never modified. Its front matter is kept as is: cover
page, introductory and legal sections, lists of figures / tables, and the
table of contents. Everything after the table of contents (the actual
content, up to the final section properties) is replaced by the two marker
paragraphs expected by md_to_docx. An optional block located before the table
of contents (typically a static "Document control" section that is now
maintained in Markdown) can be dropped too.

Image relationships that are no longer referenced by the trimmed body are
removed, together with the media files no other part of the package uses, so
the host stays small.

All edits are made on the XML tree (no regex surgery on the OOXML), and the
original namespace declarations of the document root are preserved verbatim.

Usage:
    python3 make_host_template.py \
        --source path/to/existing-document.docx \
        --out    path/to/host-template.docx \
        [--drop-block "Document control" "List of Figures"]
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
W = "{%s}" % W_NS
START_MARKER = "<!-- INSERT-START -->"
END_MARKER = "<!-- INSERT-END -->"
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
ROOT_TAG_RE = re.compile(r"<w:document\b[^>]*>")
XMLNS_RE = re.compile(r'xmlns:([A-Za-z0-9_.-]+)="([^"]+)"')


# --------------------------------------------------------------------------- #
# XML helpers (prefix-preserving round trip)
# --------------------------------------------------------------------------- #


def _register_namespaces(raw: str, root_tag: str) -> str:
    """Register every prefix declared anywhere in the part (Word declares
    some of them locally, e.g. on drawings) and return a root start tag that
    declares them all, so the serialized tree keeps the original prefixes."""
    root_decls = dict(XMLNS_RE.findall(root_tag))
    extra = []
    for prefix, uri in XMLNS_RE.findall(raw):
        if prefix == "xml":
            continue
        ET.register_namespace(prefix, uri)
        if prefix not in root_decls:
            root_decls[prefix] = uri
            extra.append(f' xmlns:{prefix}="{uri}"')
    if not extra:
        return root_tag
    return root_tag[:-1] + "".join(extra) + ">"


def _parse(path: Path) -> tuple[ET.Element, str]:
    raw = path.read_text(encoding="utf-8")
    m = ROOT_TAG_RE.search(raw)
    if not m:
        raise ValueError(f"no <w:document> root tag in {path}")
    root_tag = _register_namespaces(raw, m.group(0))
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
    root = ET.fromstring(raw, parser=parser)
    return root, root_tag


def _serialize(root: ET.Element, root_tag: str) -> str:
    body = ET.tostring(root, encoding="unicode")
    m = ROOT_TAG_RE.search(body)
    if not m:
        raise ValueError("serialized root tag not found")
    # Put back the original root start tag so every namespace declaration
    # (including the ones only referenced from mc:Ignorable) survives.
    return XML_DECL + root_tag + body[m.end():]


def _text(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.iter(W + "t")).strip()


def _is_toc_sdt(el: ET.Element) -> bool:
    if el.tag != W + "sdt":
        return False
    gallery = el.find(".//" + W + "docPartGallery")
    return gallery is not None and gallery.get(W + "val") == "Table of Contents"


def _marker_paragraph(marker: str) -> ET.Element:
    """<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>MARKER</w:t></w:r></w:p>

    The marker is plain visible text (not an XML comment): Word keeps it
    when the host is edited and re-saved, and the converter accepts both forms.
    """
    p = ET.Element(W + "p")
    ppr = ET.SubElement(p, W + "pPr")
    ET.SubElement(ppr, W + "pStyle", {W + "val": "BodyText"})
    r = ET.SubElement(p, W + "r")
    t = ET.SubElement(r, W + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
    t.text = marker
    return p


def _has_page_break(el: ET.Element) -> bool:
    return any(br.get(W + "type") == "page" for br in el.iter(W + "br"))


# --------------------------------------------------------------------------- #
# Body trimming
# --------------------------------------------------------------------------- #


def trim_body(root: ET.Element, drop_block: tuple[str, str] | None) -> dict:
    body = root.find(W + "body")
    if body is None:
        raise ValueError("document has no w:body")
    children = list(body)
    if not children or children[-1].tag != W + "sectPr":
        raise ValueError("expected the body to end with w:sectPr")
    final_sectpr = children[-1]

    toc_idx = next((i for i, el in enumerate(children) if _is_toc_sdt(el)), None)
    if toc_idx is None:
        raise ValueError("no Table of Contents content control found in the body")

    # Keep the empty paragraphs that follow the TOC up to the last one that
    # carries a page break (so the content starts on a fresh page, without
    # blank lines above it); leftover empty paragraphs and any former marker
    # paragraphs are dropped.
    keep_until = toc_idx + 1
    last_break = toc_idx
    while keep_until < len(children) - 1 and children[keep_until].tag == W + "p" \
            and not _text(children[keep_until]) and "INSERT-" not in _text(children[keep_until]):
        if _has_page_break(children[keep_until]):
            last_break = keep_until
        keep_until += 1
    if last_break > toc_idx:
        keep_until = last_break + 1

    # Former marker paragraphs (a host being rebuilt) are dropped as well.
    kept = [el for el in children[:keep_until]
            if not (el.tag == W + "p" and _text(el) in (START_MARKER, END_MARKER))]
    stats = {"dropped_content": len(children) - 1 - keep_until, "dropped_block": 0}

    if drop_block:
        start_text, end_text = drop_block
        s = next((i for i, el in enumerate(kept) if el.tag == W + "p" and _text(el) == start_text), None)
        e = next((i for i, el in enumerate(kept) if el.tag == W + "p" and _text(el) == end_text), None)
        if s is None or e is None or e <= s:
            raise ValueError(f"drop block not found: {start_text!r} .. {end_text!r}")
        # Keep the trailing page-break paragraph right before the end heading.
        if e - 1 > s and not _text(kept[e - 1]):
            e -= 1
        stats["dropped_block"] = e - s
        kept = kept[:s] + kept[e:]

    for el in children:
        body.remove(el)
    for el in kept:
        body.append(el)
    body.append(_marker_paragraph(START_MARKER))
    body.append(_marker_paragraph(END_MARKER))
    body.append(final_sectpr)
    return stats


# --------------------------------------------------------------------------- #
# Relationship / media pruning
# --------------------------------------------------------------------------- #


def _referenced_rids(root: ET.Element) -> set[str]:
    rids = set()
    for el in root.iter():
        for attr in ("{%s}id" % R_NS, "{%s}embed" % R_NS, "{%s}link" % R_NS, "{%s}pict" % R_NS):
            v = el.get(attr)
            if v:
                rids.add(v)
    return rids


def prune_media(tmp: Path, doc_root: ET.Element) -> dict:
    ET.register_namespace("", PKG_REL_NS)
    rels_path = tmp / "word" / "_rels" / "document.xml.rels"
    rels_root = ET.fromstring(rels_path.read_text(encoding="utf-8"))
    used = _referenced_rids(doc_root)
    removed = 0
    for rel in list(rels_root):
        if rel.get("Type") == IMAGE_REL_TYPE and rel.get("Id") not in used:
            rels_root.remove(rel)
            removed += 1
    rels_path.write_text(XML_DECL + ET.tostring(rels_root, encoding="unicode"), encoding="utf-8")

    # Media still referenced by any .rels file of the package are kept.
    still_used: set[str] = set()
    for rels in tmp.rglob("*.rels"):
        r = ET.fromstring(rels.read_text(encoding="utf-8"))
        for rel in r:
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External":
                continue
            still_used.add((rels.parent.parent / target).resolve().as_posix()
                           if not target.startswith("/") else (tmp / target.lstrip("/")).resolve().as_posix())
    deleted = 0
    media_dir = tmp / "word" / "media"
    if media_dir.exists():
        for f in sorted(media_dir.iterdir()):
            if f.resolve().as_posix() not in still_used:
                f.unlink()
                deleted += 1
    return {"image_rels_removed": removed, "media_deleted": deleted}


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #


def build(source: Path, out: Path, drop_block: tuple[str, str] | None) -> dict:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        with zipfile.ZipFile(source) as zf:
            zf.extractall(tmp)
        doc_path = tmp / "word" / "document.xml"
        root, root_tag = _parse(doc_path)
        stats = trim_body(root, drop_block)
        doc_path.write_text(_serialize(root, root_tag), encoding="utf-8")
        stats.update(prune_media(tmp, root))

        if out.exists():
            out.unlink()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml first, as Word writes it.
            ct = tmp / "[Content_Types].xml"
            zf.write(ct, "[Content_Types].xml")
            for p in sorted(tmp.rglob("*")):
                if p.is_file() and p != ct:
                    zf.write(p, p.relative_to(tmp).as_posix())
    return stats


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--source", required=True, type=Path, help="fully written .docx to derive the host from (read only)")
    ap.add_argument("--out", required=True, type=Path, help="host template .docx to write")
    ap.add_argument("--drop-block", nargs=2, metavar=("START_HEADING", "END_HEADING"),
                    default=("Document control", "List of Figures"),
                    help="drop the front-matter block from START_HEADING (inclusive) to END_HEADING (exclusive); "
                         "default: 'Document control' .. 'List of Figures'")
    ap.add_argument("--keep-all-front-matter", action="store_true", help="disable --drop-block")
    args = ap.parse_args()
    if not args.source.exists():
        sys.exit(f"source not found: {args.source}")
    if args.source.resolve() == args.out.resolve():
        sys.exit("refusing to overwrite the source document")
    stats = build(args.source, args.out, None if args.keep_all_front_matter else tuple(args.drop_block))
    print(f"written: {args.out}")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
