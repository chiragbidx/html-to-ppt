# html-to-ppt REST API

Simple Node.js REST API that converts raw HTML into a native `.pptx` file.

## Run

```bash
npm install
npx playwright install chromium
npm start
```

Server runs on `http://localhost:3000` by default.

## Endpoints

- `GET /health`
- `POST /api/convert`
- `POST /api/convert-raw`

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
- `browser` (default): full HTML/CSS rendering fidelity by screenshotting rendered slide(s)
- `simple`: basic native text extraction (`h1`-`h3`, `p`, `ul`, `ol`, `li`)

### Example cURL

```bash
curl -X POST http://localhost:3000/api/convert \
  -H "Content-Type: application/json" \
  -d '{
    "title": "Demo",
    "filename": "demo-presentation",
    "renderMode": "browser",
    "html": "<h1>Hello</h1><p>This PPT was created from HTML.</p><ul><li>Point one</li><li>Point two</li></ul>"
  }' \
  --output demo-presentation.pptx
```

### Raw HTML (no JSON escaping)

Send HTML directly as request body with `Content-Type: text/html`. Optional `title`, `filename`, and `renderMode` can be passed as query params.

```bash
curl -X POST "http://localhost:3000/api/convert-raw?title=Demo%20Raw&filename=demo-raw&renderMode=browser" \
  -H "Content-Type: text/html" \
  --data '<h1>Hello</h1><p>This was sent as raw HTML.</p><ul><li>Point one</li><li>Point two</li></ul>' \
  --output demo-raw.pptx
```

## Notes

- For CSS-heavy decks, structure each page inside `.slide` containers sized around `1280x720`. Each `.slide` becomes one PowerPoint slide in `browser` mode.
- Browser mode preserves visual output, but text is embedded as slide images (not directly editable text in PPT).
