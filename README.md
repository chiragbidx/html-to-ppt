# html-to-ppt REST API

Simple Node.js REST API that converts raw HTML into a `.pptx` file.

## Run

```bash
npm install
npx playwright install chromium
npm start
```

Server runs on `http://localhost:3000` by default.

## Docker

```bash
docker build -t html-to-ppt .
docker run --rm -p 3000:3000 -e PORT=3000 html-to-ppt
```

## Endpoints

- `GET /health`
- `POST /api/convert`
- `POST /api/convert-raw`
- `POST /api/convert-pdf`

### Request body

```json
{
  "title": "Quarterly Update",
  "filename": "quarterly-update",
  "renderMode": "browser",
  "html": "<h1>Quarterly Update</h1><p>Revenue up 17% YoY.</p><ul><li>New enterprise customers</li><li>Expanded product line</li></ul>"
}
```

`renderMode` values:
- `browser` (default): full HTML/CSS rendering fidelity by screenshotting rendered slide(s) into PPT image slides.
- `simple`: native editable PPT rendering using browser-computed layout. Supports text styling (`color`, `font-size`, `font-weight`, `font-style`, `text-align`, `line-height`, `letter-spacing`, `text-transform`) and box styles (`background`, `border`, `border-radius`) for many HTML structures.

### Example cURL

```bash
curl -X POST http://localhost:3000/api/convert \
  -H "Content-Type: application/json" \
  -d '{
    "title": "Demo",
    "filename": "demo-presentation",
    "renderMode": "simple",
    "html": "<h1>Hello</h1><p>This PPT was created from HTML.</p><ul><li>Point one</li><li>Point two</li></ul>"
  }' \
  --output demo-presentation.pptx
```

### Raw HTML (no JSON escaping)

Send HTML directly as request body with `Content-Type: text/html`. Optional `title`, `filename`, and `renderMode` can be passed as query params.

```bash
curl -X POST "http://localhost:3000/api/convert-raw?title=Demo%20Raw&filename=demo-raw&renderMode=simple" \
  -H "Content-Type: text/html" \
  --data '<h1>Hello</h1><p>This was sent as raw HTML.</p><ul><li>Point one</li><li>Point two</li></ul>' \
  --output demo-raw.pptx
```

### PDF upload to PPT

Send a PDF as `multipart/form-data` field `file`. Optional form fields: `title`, `filename`, `pageRange`, `renderMode`.

```bash
curl -X POST http://localhost:3000/api/convert-pdf \
  -F "file=@./report.pdf;type=application/pdf" \
  -F "title=PDF Demo" \
  -F "filename=pdf-demo" \
  -F "renderMode=browser" \
  -F "pageRange=1-3,5" \
  --output pdf-demo.pptx
```

Response headers include:
- `X-Editable-Pages`: pages converted to editable PPT text/shapes.
- `X-Fallback-Pages`: pages rendered as image fallback.
- `X-PDF-Render-Mode`: resolved mode (`simple` or `browser`).

## Notes

- For slide-separated input, wrap each page in `.slide` containers around `1280x720`. In both modes, each `.slide` maps to one PowerPoint slide.
- `simple` mode preserves browser-calculated placement from layouts like grid/flex by mapping rendered element bounds into native PPT text boxes and shapes.
- `simple` mode still does not support full browser parity for effects such as shadows, transforms, filters, gradients, and canvas/SVG rendering semantics.
- `browser` mode preserves visual output via screenshots, but text is image-based and not directly editable.
- `convert-pdf` supports `renderMode=simple` (editable-first) and `renderMode=browser` (PDF page images).
- In `simple` mode, if per-page fidelity checks fail, the page is inserted as an image-backed slide.
