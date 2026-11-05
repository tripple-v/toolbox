# puml-to-image

Renders PlantUML source files (`.puml`) to PNG or SVG images on your own machine, without any online service.

## Use

```bash
./render_puml.sh diagram.puml                  # diagram.png and diagram.svg next to the source
./render_puml.sh -f png diagram.puml           # PNG only
./render_puml.sh -o images/ a.puml b.puml      # several files, images in images/
```

The script prints the path of each image it wrote.

## What you need

- `plantuml` on the PATH. On macOS: `brew install plantuml` (it brings Java and Graphviz with it). On Windows: `choco install plantuml` from a PowerShell window opened as administrator (same dependencies brought in), then run the script from Git Bash. Elsewhere, see https://plantuml.com/download: a Java runtime and the `plantuml.jar`, wrapped in a `plantuml` command.
- Tested with PlantUML 1.2026.6.

## How it works

The script calls `plantuml -t<format> -o <output dir> <file>` once per file and per format. The image takes the name of the source file; if the source declares a name after `@startuml`, PlantUML uses that name instead. Everything runs locally: the diagram never leaves the machine, which matters for diagrams that describe internal systems.

## Good to know

- A sequence diagram with many participants gets wide; prefer SVG for reading on screen and PNG for pasting into a document.
- Fonts: PlantUML uses the fonts installed on the machine. To fix the font, declare it in the source (`skinparam defaultFontName Arial`).
- Errors in the source are reported by PlantUML on stderr with the line number; the script stops at the first failing file.
