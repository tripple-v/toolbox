# openapi-to-html

Turns an OpenAPI file into a single HTML page that opens in any browser, without a network connection. Handy to send an API contract to someone who has no OpenAPI tool at hand.

## Use

```bash
python3 build_openapi_html.py path/to/api.openapi.yaml [output.html]
```

Without a second argument, the page is written next to the OpenAPI file, with the `.html` extension.

## What you need

- **Python 3.9 or later**, no extra package
- **Node.js**, for the `npx` command
- **a network connection while the page is built**

## How it works

1. The Redocly command-line tool renders the OpenAPI file as an HTML page. It is run through `npx`, which downloads it on first use.
2. That page loads its viewer script and its fonts from the internet. The script is downloaded once and embedded in the page, and the link to the online fonts is removed.

The OpenAPI file is processed on your machine and is not uploaded anywhere. The resulting page needs no connection.

## Good to know

- Rebuild the page whenever the OpenAPI file changes: the page is a snapshot.
- The latest version of the rendering tool is used each time, so the look of the page can change slightly from one build to the next.
- If you do not need the page to work offline, the first step alone is enough: `npx @redocly/cli build-docs api.openapi.yaml -o page.html`.
