# Toolbox

Small, self-contained tools for everyday technical work. Each tool lives in its own folder, with its own README, and can be copied and used on its own.

## Tools

| Tool | What it does |
| --- | --- |
| [md-to-docx](./md-to-docx/README.md) | Builds a Word document from Markdown files, with the look of an existing Word template |
| [openapi-to-html](./openapi-to-html/README.md) | Turns an OpenAPI file into a single HTML page that opens in any browser, without a network connection |
| [puml-to-image](./puml-to-image/README.md) | Renders PlantUML files to PNG or SVG images locally |

## Adding a tool

1. Create a folder named after what the tool does, in lower case with hyphens (`openapi-to-html`).
2. Put everything the tool needs in that folder. A tool does not depend on another tool.
3. Write a `README.md` in the folder, with these sections:
   - what the tool does, in one or two sentences;
   - **Use**: the command to run, or the steps to follow;
   - **What you need**: other tools and versions;
   - **How it works**: enough for a reader to trust it and to fix it;
   - **Good to know**: limits and pitfalls.
4. Add a line to the table above.

Keep a tool generic: no project name, no customer name, no credentials, no internal address. If a tool needs such a value, make it a parameter or a configuration file kept outside this repository.

## License

Released under the [MIT License](./LICENSE).
