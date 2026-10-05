# md-to-docx

Builds a Word document (`.docx`) from Markdown files, with the look of an existing Word template.

You keep writing in Markdown. The tool takes a Word file that already has your cover page, table of contents, headers, footers and styles (the *host template*), and fills it with the content of your Markdown files.

## What is in this folder

| File                    | What it does                                                            |
|-------------------------|-------------------------------------------------------------------------|
| `build_docx.py`         | The command you run. Reads a configuration file and builds the document |
| `md_to_docx.py`         | The converter used by `build_docx.py`. You rarely call it yourself      |
| `make_host_template.py` | Turns an existing Word document into a host template                    |
| `examples/`             | A host template and two sample documents, ready to build                |

## 1. Install

You need:

- **Python 3.11 or later**
- the Python packages listed in `requirements.txt`
- **rsvg-convert**, only if your Markdown embeds SVG images (they are converted to PNG)

```bash
pip install -r requirements.txt
```

To install rsvg-convert: `brew install librsvg` on macOS, `apt install librsvg2-bin` on Debian or Ubuntu, or the librsvg installer on Windows.

## 2. Try the examples

From the current folder:

```bash
python3 build_docx.py examples/single/sample.json
python3 build_docx.py examples/multi/document.json
```

Each command prints the path of the document it wrote, under `examples/.../output/`.

Open the result in Word. Word asks: *"This document contains fields that may refer to other files. Do you want to update the fields in this document?"* Answer **Yes**: this fills in the table of contents, the figure numbers and the page count. Then save the document.

Compare `examples/single/sample.md` with the document you get: the sample uses every feature described below.

## 3. Build your own document

You need three things: your Markdown, a host template, and a configuration file.

### Step 1 - Prepare a host template

A host template is a normal Word document containing two paragraphs, each on its own line, exactly as written here:

```
<!-- INSERT-START -->
<!-- INSERT-END -->
```

The content generated from your Markdown replaces these two paragraphs. Everything else in the document is kept: cover page, table of contents, legal sections, headers, footers.

Two ways to get one:

- **By hand.** Open your corporate template in Word, write your cover page, insert a table of contents, then type the two marker lines where the content must go. Save as `.docx`.
- **From an existing document.** If you already have a complete document written with the template:

  ```bash
  python3 make_host_template.py --source existing-document.docx --out host-template.docx
  ```

  Everything after the table of contents is replaced by the two markers. The source document is not modified.

### Step 2 - Tell the tool which styles your template uses

The tool applies a Word style to each kind of content. By default it uses the standard Word style identifiers:

| Content                              | Default style identifier |
|--------------------------------------|--------------------------|
| Headings, levels 1 to 6              | `Heading1` to `Heading6` |
| Front-matter headings (not numbered) | `Title`, then `Subtitle` |
| Paragraph                            | `BodyText`               |
| Bullet list                          | `ListBullet`             |
| Numbered list                        | `ListNumber`             |
| Caption                              | `Caption`                |
| Table                                | `TableGrid`              |

If your template uses other identifiers, write them in a small JSON file and name it in your configuration (key `styles`). You only list what differs:

```json
{
  "frontmatter_heading": { "1": "Heading7", "2": "Heading5", "3": "Heading6" },
  "table": "MyCompanyTable",
  "content_width_twips": 9620
}
```

`content_width_twips` is the width of the text area, in twentieths of a point: page width minus the left and right margins. It is used to size tables and to scale down large images. The default suits an A4 page with margins of about 2 cm.

How to find a style identifier: in Word it is usually the style name without spaces (the style "Body Text" has the identifier `BodyText`). If a style does not apply, check the name in the Word styles pane.

### Step 3 - Write the configuration file

A configuration file is a JSON file. Paths inside it are relative to the configuration file itself, so keep it next to your documents.

One Markdown file:

```json
{
  "source": "my-document.md",
  "host": "host-template.docx",
  "out": "output/my-document.docx"
}
```

Several Markdown files, merged in order into one document:

```json
{
  "sources": [
    "01-introduction.md",
    "02-design.md",
    "03-operations.md"
  ],
  "host": "host-template.docx",
  "out": "output/my-document.docx",
  "strip_heading_numbers": true
}
```

### Step 4 - Build

```bash
python3 build_docx.py path/to/my-config.json
```

Open the result in Word, answer **Yes** to the field update question, and save.

## 4. Configuration reference

### Required

| Key | Meaning |
| --- | --- |
| `source` or `sources` | One Markdown file, or a list of files merged in order |
| `host` | The host template. May contain `*`: the matching file with the highest version number in its name is used, for example `templates/host_v*.docx` |
| `out` | The document to write. May contain `{version}`, see `version_row` |

### Optional, for the whole document

| Key | Meaning |
| --- | --- |
| `styles` | JSON file with the style identifiers of your template (step 2) |
| `body_starts_at` | Title of the first heading of the body. Every heading before it is treated as front matter: not numbered, no page break. Use it when the document starts with a title and a document control block |
| `version_row` | Label of the table row that holds the version in the first source. Default `Version`: a row `\| Version \| 1.0 \|` turns `{version}` into `v1_0` in the output name |
| `keep_internal_blocks` | `true` to keep the blocks marked as internal (see the markers below). By default they are left out |
| `work_dir` | Folder for the temporary files. Default: a folder next to the output |

### Optional, for every source

These apply to all sources. A source can override them (see below).

| Key | Meaning |
| --- | --- |
| `drop_title` | `true` to drop the first level-1 heading of the source, when the title is already on the cover page |
| `strip_heading_numbers` | `true` to remove numbers typed by hand in headings (`3. Title`), when the template numbers the headings itself |
| `heading_shift` | Number of levels added to every heading. `1` turns `#` into `##`; `-1` turns `##` into `#` |
| `drop_note_after_title` | `true` to drop the blockquote that directly follows the first heading, for example a navigation note |
| `drop_sections` | List of heading titles. Each section is removed with its content |
| `flatten_links` | `true` to replace links by their text. Useful when the links point to other files of a repository |
| `plain_checkboxes` | `true` to turn `- [ ] item` into `- item` |
| `replace_lines` | List of `[pattern, replacement]` pairs (regular expressions) applied to each line |

### Per source

In `sources`, an entry can be an object instead of a path. It then carries its own options, plus `insert_before`:

```json
{
  "sources": [
    "01-context.md",
    { "path": "02-design-overview.md", "insert_before": "# Design", "drop_title": true },
    { "path": "02-design-details.md", "drop_title": true }
  ]
}
```

Here the two design files lose their own titles and sit under a single chapter heading, "Design". This is the `examples/multi` sample.

### Several editions of one document

To produce, for example, an internal and an external edition from the same Markdown, write two configuration files. They point to the same sources and differ by `out`, `drop_sections` and `keep_internal_blocks`. The Markdown is not touched.

## 5. What you can write in Markdown

| Markdown | Result in Word |
| --- | --- |
| `# Heading` to `###### Heading` | Headings, numbered by the template |
| Paragraphs, `*italic*`, `**bold**` | Body text |
| `- item` | Bullet list |
| `1. item` | Numbered list |
| A pipe table with a header row | Table, header row in bold |
| `![alt](path/to/image.png)` | Image, scaled down to the page width if needed. PNG, JPEG and SVG |
| `![alt](image.png){width=60%}` | Image at 60 % of the page width |
| `*Figure: text*` alone in a paragraph, after an image | Caption "Figure 3: text", listed in the list of figures |
| `*Table: text*` alone in a paragraph, after a table | Caption "Table 2: text" |
| `[Auto]` alone in a table cell | Total number of pages |
| `<br>` inside a paragraph or a table cell | Line break |
| `` `code` `` and fenced code blocks | Code style, if the template has one (see below) |

A new page starts automatically before each first-level chapter.

### Layout markers

These are HTML comments. They do not appear when the Markdown is rendered on a web page. Write each on a line of its own, with an empty line before and after.

| Marker | Effect |
| --- | --- |
| `<!-- pagebreak -->` | Starts a new page |
| `<!-- linebreak -->` | Inserts an empty line |
| `<!-- colwidths: 20 50 30 -->` | Sets the column widths of the table that follows. One value per column; the values are proportions |
| `<!-- internal:start -->` ... `<!-- internal:end -->` | The lines in between are left out of the Word document |

Example:

```markdown
<!-- colwidths: 20 80 -->

| Term | Definition |
| --- | --- |
| Host template | The Word file that carries the cover page and the styles |

<!-- internal:start -->

## Notes for the authors

Not exported.

<!-- internal:end -->
```

A block marked as internal is only hidden from the Word document. Anyone who can read the Markdown file can still read it.

### Code

If the host template has a paragraph style named `Code`, code blocks use it. If it has a character style named `Code Char`, inline code uses it. Without these styles, code is written as normal text.

To create the style in Word: open the host template, open the styles pane, choose *New Style*. Name it `Code`, type *Linked (paragraph and character)*, based on *Normal*, with a fixed-width font such as Consolas or Courier New. In the paragraph format, set the spacing before and after to 0 and tick "Don't add space between paragraphs of the same style". Word then creates `Code Char` by itself. Save the template.

## 6. Troubleshooting

| Symptom | Cause and fix |
| --- | --- |
| `START_MARKER ... not found in host document.xml` | The host template has no `<!-- INSERT-START -->` paragraph. Type it in Word, on a line of its own |
| The table of contents is empty or out of date | Open the document in Word and answer **Yes** to the field update question, or select all and press F9 |
| Text is not formatted as expected | The style identifier does not exist in the template. Declare the right one in the styles file (step 2) |
| `images not found` | A path in `![...](...)` is wrong. Image paths are relative to the Markdown file that contains them |
| `rsvg-convert` not found | Install it (section 1), or use PNG or JPEG images |
| `warning: colwidths gives 2 value(s) for a table of 3 column(s)` | The marker must give one value per column. The table is built with equal columns |
| `warning: internal block ... is never closed` | An `<!-- internal:start -->` has no matching `<!-- internal:end -->`. Everything after it was left out |
| A marker appears as text in the document | It is not on a line of its own, or it has no empty line before and after |
| Headings are numbered twice | The Markdown has numbers typed by hand and the template numbers headings too. Set `strip_heading_numbers` |

## 7. Not supported

- merged cells in tables
- lists nested inside lists, lists or tables inside table cells
- footnotes
- HTML, apart from `<br>` and the markers above
