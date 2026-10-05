#!/usr/bin/env python3
"""Build a styled .docx from one or several Markdown files, driven by a JSON
configuration file.

    python3 build_docx.py path/to/config.json [--out other-output.docx]

The Markdown sources are never modified. A working copy is prepared (sources
merged in order, headings adjusted, images staged), then handed to the
converter (md_to_docx.py) together with the host template.

Paths in the configuration are relative to the configuration file.

Required keys
    source                  one Markdown file, or
    sources                 a list of Markdown files, merged in that order. An
                            entry is a path, or an object with a "path" and
                            options of its own (see "Per-source options")
    host                    host template .docx holding the two insert markers.
                            May contain "*": the matching file with the highest
                            version number in its name is used
    out                     output .docx. May contain "{version}", replaced by
                            the version read in the first source (see
                            "version_row")

Optional keys
    styles                  JSON file giving the style ids of the host template
                            (passed to the converter with --styles)
    work_dir                where the working copy and the images are staged
                            (default: "<output name>_work" next to the output)
    converter               converter script (default: md_to_docx.py next to
                            this file)
    body_starts_at          title of the first heading of the body; every
                            heading before it is marked as front matter (not
                            numbered, no page break), for example a title and
                            a document control block
    version_row             label of the table row holding the document version
                            in the first source (default "Version"). A row
                            "| Version | 1.0 |" gives {version} = "v1_0"
    keep_internal_blocks    true to keep the blocks enclosed in the sources by
                            the internal:start / internal:end markers; by
                            default they are left out of the document

Options applying to every source, unless a source overrides them
    drop_title              true to drop the first level-1 heading of the
                            source (when the title is on the cover page of the
                            host, or replaced by "insert_before")
    strip_heading_numbers   true to remove manual numbers ("3. Title"), when
                            the host template numbers the headings itself
    heading_shift           number of levels added to every heading: 1 turns
                            "#" into "##", -1 turns "##" into "#" (default 0)
    drop_note_after_title   true to drop the blockquote that directly follows
                            the first heading (for example a navigation note)
    drop_sections           list of heading titles; each section is removed
                            with its content, up to the next heading of the
                            same or a higher level. Several configuration files
                            can thus produce several editions of one document
    flatten_links           true to replace links by their text (images are
                            kept)
    plain_checkboxes        true to turn "- [ ] item" into "- item"
    replace_lines           list of [pattern, replacement] pairs (regular
                            expressions) applied to each line outside code
                            blocks

Per-source options
    Any of the options above, plus:
    insert_before           Markdown line inserted before the source, for
                            example "# Use case view" to group two sources
                            under one chapter heading

Requires rsvg-convert on the PATH when the sources embed SVG images.
"""

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote

HEAD = re.compile(r"^(#{1,6})\s+(.*)$")
NUMBER = re.compile(r"^\d+(?:\.\d+)*\.?\s+")
IMAGE = re.compile(r"(!\[[^\]]*\]\()([^)]+)(\))")
LINK = re.compile(r"(?<!!)\[([^\]]+)\]\([^)]+\)")
CHECKBOX = re.compile(r"^(\s*)- \[[ xX]\] ")
PER_SOURCE = ("drop_title", "strip_heading_numbers", "heading_shift", "drop_note_after_title",
              "drop_sections", "flatten_links", "plain_checkboxes", "replace_lines")


def pick_host(base: Path, pattern: str) -> Path:
    """The host file, or for a pattern with "*" the match with the highest
    version (the numbers found in the file name, compared in order)."""
    if "*" not in pattern:
        return (base / pattern).resolve()
    target = base / pattern
    matches = [p for p in target.parent.glob(target.name) if not p.name.startswith("~$")]
    if not matches:
        sys.exit(f"no host template matches: {target}")
    return max(matches, key=lambda p: [int(n) for n in re.findall(r"\d+", p.stem)]).resolve()


def read_version(text: str, row: str) -> str:
    pattern = re.compile(r"^\|\s*" + re.escape(row) + r"\s*\|\s*(\d+)\.(\d+)")
    for line in text.splitlines():
        match = pattern.match(line)
        if match:
            return f"v{match.group(1)}_{match.group(2)}"
    sys.exit(f'the output name uses {{version}} but no "| {row} | x.y |" row was found in the first source')


def drop_note_after_title(lines: list) -> list:
    out, i, state = [], 0, "before_title"
    while i < len(lines):
        line = lines[i]
        if state == "before_title" and line.startswith("# "):
            state = "after_title"
        elif state == "after_title":
            if line.strip() and line.startswith(">"):
                while i < len(lines) and (lines[i].startswith(">") or not lines[i].strip()):
                    i += 1
                state = "done"
                continue
            if line.strip():
                state = "done"
        out.append(line)
        i += 1
    return out


def drop_sections(lines: list, titles: list) -> list:
    wanted = {t.strip().lower() for t in titles}
    if not wanted:
        return lines
    out, skipping, in_code = [], 0, False
    for line in lines:
        if line.startswith("```"):
            in_code = not in_code
        match = None if in_code else HEAD.match(line)
        if match:
            level = len(match.group(1))
            if skipping and level <= skipping:
                skipping = 0
            if NUMBER.sub("", match.group(2)).strip().lower() in wanted:
                skipping = level
        if not skipping:
            out.append(line)
    return out


class Stager:
    """Copies the images referenced by the sources into one flat folder,
    converting SVG to PNG, and gives each a unique name."""

    def __init__(self, images_dir: Path):
        self.images_dir = images_dir
        self.staged: dict = {}
        self.missing: list = []

    def stage(self, src_dir: Path, ref: str) -> str:
        src = (src_dir / unquote(ref)).resolve()
        if not src.exists():
            self.missing.append(str(src))
            return ref
        if src in self.staged:
            return self.staged[src]
        suffix = ".png" if src.suffix.lower() == ".svg" else src.suffix
        name, n = src.stem + suffix, 1
        while (self.images_dir / name).exists():
            n += 1
            name = f"{src.stem}_{n}{suffix}"
        dest = self.images_dir / name
        if src.suffix.lower() == ".svg":
            subprocess.run(["rsvg-convert", "-o", str(dest), str(src)], check=True)
        else:
            shutil.copy2(src, dest)
        self.staged[src] = f"images/{name}"
        return self.staged[src]


class Front:
    """Tracks whether the headings met so far belong to the front matter."""

    def __init__(self, body_starts_at):
        self.target = (body_starts_at or "").strip().lower()
        self.active = bool(self.target)

    def is_front(self, title: str) -> bool:
        if self.active and NUMBER.sub("", title).strip().lower() == self.target:
            self.active = False
        return self.active


def prepare(text: str, opts: dict, src_dir: Path, stager: Stager, front: Front) -> str:
    lines = text.splitlines()
    if opts.get("drop_note_after_title"):
        lines = drop_note_after_title(lines)
    lines = drop_sections(lines, opts.get("drop_sections") or [])
    replacements = [(re.compile(p), r) for p, r in opts.get("replace_lines") or []]
    shift = int(opts.get("heading_shift") or 0)
    out, in_code, title_dropped = [], False, False
    for line in lines:
        if line.startswith("```"):
            in_code = not in_code
        if not in_code:
            match = HEAD.match(line)
            if match:
                level, title = len(match.group(1)), match.group(2)
                if opts.get("drop_title") and level == 1 and not title_dropped:
                    title_dropped = True
                    continue
                is_front = front.is_front(title)
                if opts.get("strip_heading_numbers"):
                    title = NUMBER.sub("", title)
                line = "#" * min(6, max(1, level + shift)) + " " + title.rstrip()
                if is_front:
                    line += " {.frontmatter}"
            line = IMAGE.sub(lambda m: m.group(1) + stager.stage(src_dir, m.group(2)) + m.group(3), line)
            if opts.get("flatten_links"):
                line = LINK.sub(lambda m: m.group(1), line)
            if opts.get("plain_checkboxes"):
                line = CHECKBOX.sub(r"\1- ", line)
            for pattern, replacement in replacements:
                line = pattern.sub(replacement, line)
        out.append(line)
    return "\n".join(out)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a styled .docx from Markdown, from a JSON configuration file.")
    parser.add_argument("config", type=Path, help="configuration file (JSON)")
    parser.add_argument("--out", type=Path, help="output .docx, instead of the one named in the configuration")
    args = parser.parse_args()

    cfg_path = args.config.resolve()
    if not cfg_path.exists():
        sys.exit(f"configuration not found: {cfg_path}")
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    base = cfg_path.parent

    entries = cfg.get("sources") or ([cfg["source"]] if cfg.get("source") else [])
    if not entries:
        sys.exit('the configuration names no "source" nor "sources"')
    for key in ("host", "out"):
        if key not in cfg:
            sys.exit(f"missing configuration key: {key}")
    sources = []
    for entry in entries:
        opts = {key: cfg[key] for key in PER_SOURCE if key in cfg}
        if isinstance(entry, dict):
            opts.update({k: v for k, v in entry.items() if k != "path"})
            entry = entry["path"]
        path = (base / entry).resolve()
        if not path.exists():
            sys.exit(f"source not found: {path}")
        sources.append((path, opts))

    host = pick_host(base, cfg["host"])
    if not host.exists():
        sys.exit(f"host template not found: {host}")
    first_text = sources[0][0].read_text(encoding="utf-8")
    out_name = str(args.out) if args.out else cfg["out"]
    if "{version}" in out_name:
        out_name = out_name.replace("{version}", read_version(first_text, cfg.get("version_row", "Version")))
    out = Path(out_name).resolve() if args.out else (base / out_name).resolve()

    tool_dir = Path(__file__).resolve().parent
    converter = (base / cfg["converter"]).resolve() if cfg.get("converter") else tool_dir / "md_to_docx.py"
    work = (base / cfg["work_dir"]).resolve() if cfg.get("work_dir") else out.with_name(out.stem + "_work")
    images_dir = work / "images"
    if images_dir.exists():
        shutil.rmtree(images_dir)
    images_dir.mkdir(parents=True)

    stager, front, parts = Stager(images_dir), Front(cfg.get("body_starts_at")), []
    for path, opts in sources:
        if opts.get("insert_before"):
            parts.append(opts["insert_before"] + "\n")
        parts.append(prepare(path.read_text(encoding="utf-8"), opts, path.parent, stager, front))
    if stager.missing:
        sys.exit("images not found:\n  " + "\n  ".join(stager.missing))
    merged = work / "merged.md"
    merged.write_text("\n\n".join(parts) + "\n", encoding="utf-8")

    out.parent.mkdir(parents=True, exist_ok=True)
    command = [sys.executable, str(converter), "--md", str(merged), "--host-docx", str(host), "--out", str(out)]
    if cfg.get("styles"):
        command += ["--styles", str((base / cfg["styles"]).resolve())]
    if cfg.get("keep_internal_blocks"):
        command.append("--keep-internal")
    result = subprocess.run(command, capture_output=True, text=True)
    sys.stdout.write(result.stdout)
    sys.stderr.write(result.stderr)
    if result.returncode != 0:
        return result.returncode
    print(f"sources: {len(sources)}")
    print(f"host:    {host}")
    print(f"written: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
