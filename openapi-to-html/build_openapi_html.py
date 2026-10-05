#!/usr/bin/env python3
"""Build a single, self-contained HTML page from an OpenAPI document, readable
in any browser without a network connection.

    python3 build_openapi_html.py path/to/contract.openapi.yaml [output.html]

Two steps:
  1. the Redocly command-line tool renders the OpenAPI document as an HTML page
     (run through npx, so Node.js must be installed; the tool is downloaded on
     first use);
  2. that page still loads its viewer script and its fonts from the internet,
     so the script is downloaded once and embedded in the page, and the link
     to the online fonts is removed.

Building needs a network connection; the resulting page does not.
"""

import re
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

SCRIPT_TAG = re.compile(r'<script[^>]*src="(https?://[^"]+)"[^>]*>\s*</script>')
FONT_LINK = re.compile(r'<link[^>]*href="https?://fonts\.[^"]*"[^>]*>')
EXTERNAL = re.compile(r'<(?:script|link)[^>]*(?:src|href)="(https?://[^"]+)"')


def main() -> int:
    if len(sys.argv) not in (2, 3):
        sys.exit(__doc__)
    source = Path(sys.argv[1]).resolve()
    if not source.exists():
        sys.exit(f"not found: {source}")
    output = Path(sys.argv[2]).resolve() if len(sys.argv) == 3 else source.with_suffix(".html")

    with tempfile.TemporaryDirectory() as tmp:
        page = Path(tmp) / "page.html"
        built = subprocess.run(["npx", "--yes", "@redocly/cli@latest", "build-docs",
                                str(source), "-o", str(page)],
                               capture_output=True, text=True)
        if built.returncode != 0 or not page.exists():
            sys.stderr.write(built.stdout + built.stderr)
            sys.exit("the rendering step failed")
        html = page.read_text(encoding="utf-8")

    def embed(match: re.Match) -> str:
        with urllib.request.urlopen(match.group(1), timeout=60) as response:
            script = response.read().decode("utf-8")
        return "<script>" + script.replace("</script>", "<\\/script>") + "</script>"

    html = SCRIPT_TAG.sub(embed, html)
    html = FONT_LINK.sub("", html)
    remaining = sorted(set(EXTERNAL.findall(html)))
    if remaining:
        print("warning: the page still refers to online resources:", ", ".join(remaining),
              file=sys.stderr)
    output.write_text(html, encoding="utf-8")
    print(f"written: {output} ({len(html) // 1024} KiB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
