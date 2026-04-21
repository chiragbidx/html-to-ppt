const express = require('express');
const { htmlToPptBuffer } = require('./converter');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '5mb' }));
app.use('/api/convert-raw', express.text({ type: 'text/html', limit: '5mb' }));

app.get('/health', (_req, res) => {
  res.json({ status: 'ok' });
});

app.post('/api/convert', async (req, res) => {
  try {
    const { html, title, filename, renderMode = 'browser' } = req.body || {};

    if (!html || typeof html !== 'string') {
      return res.status(400).json({ error: 'Request body must include a non-empty string field: html' });
    }

    const buffer = await htmlToPptBuffer(html, title, renderMode);
    const safeName = (filename || 'presentation').replace(/[^a-zA-Z0-9-_]/g, '_');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}.pptx"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: 'Failed to convert HTML to PPTX', detail: error.message });
  }
});

app.post('/api/convert-raw', async (req, res) => {
  try {
    const html = req.body;
    const title = req.query.title;
    const filename = req.query.filename;
    const renderMode = req.query.renderMode || 'browser';

    if (!html || typeof html !== 'string') {
      return res
        .status(400)
        .json({ error: 'Request body must be raw HTML text with Content-Type: text/html' });
    }

    const buffer = await htmlToPptBuffer(html, title, renderMode);
    const safeName = (filename || 'presentation').replace(/[^a-zA-Z0-9-_]/g, '_');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}.pptx"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: 'Failed to convert raw HTML to PPTX', detail: error.message });
  }
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`HTML-to-PPT API listening on port ${PORT}`);
  });
}

module.exports = app;
