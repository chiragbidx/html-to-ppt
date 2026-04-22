const express = require('express');
const multer = require('multer');
const { htmlToPptBuffer } = require('./converter');
const { pdfToPptBuffer, MAX_PDF_SIZE_BYTES } = require('./pdf-converter');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '5mb' }));
app.use('/api/convert-raw', express.text({ type: 'text/html', limit: '5mb' }));

const uploadPdf = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_PDF_SIZE_BYTES },
});

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

app.post('/api/convert-pdf', (req, res) => {
  uploadPdf.single('file')(req, res, async (uploadError) => {
    if (uploadError) {
      if (uploadError.code === 'LIMIT_FILE_SIZE') {
        return res.status(400).json({ error: 'PDF file is too large', detail: uploadError.message });
      }
      return res.status(400).json({ error: 'Invalid PDF upload', detail: uploadError.message });
    }

    try {
      const file = req.file;
      const { title, filename, pageRange, renderMode = 'simple' } = req.body || {};

      if (!file || !Buffer.isBuffer(file.buffer) || file.buffer.length === 0) {
        return res.status(400).json({ error: 'Request must include multipart field: file (PDF)' });
      }

      const allowedMimeTypes = new Set(['application/pdf', 'application/octet-stream']);
      if (!allowedMimeTypes.has(file.mimetype)) {
        return res.status(400).json({ error: 'Uploaded file must be a PDF', detail: `Received: ${file.mimetype}` });
      }

      const { buffer, metadata } = await pdfToPptBuffer(file.buffer, { title, pageRange, renderMode });
      const safeName = (filename || 'presentation').replace(/[^a-zA-Z0-9-_]/g, '_');

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', `attachment; filename="${safeName}.pptx"`);
      res.setHeader('X-Editable-Pages', String(metadata.editablePages));
      res.setHeader('X-Fallback-Pages', String(metadata.fallbackPages));
      res.setHeader('X-PDF-Render-Mode', renderMode);
      return res.send(buffer);
    } catch (error) {
      if (
        /Invalid pageRange|document bounds|max pages|max size|empty|Invalid renderMode/i.test(error.message || '')
      ) {
        return res.status(400).json({ error: 'Invalid PDF conversion request', detail: error.message });
      }
      return res.status(500).json({ error: 'Failed to convert PDF to PPTX', detail: error.message });
    }
  });
});

if (require.main === module) {
  process.on('uncaughtException', (err) => {
    console.error('uncaughtException', err);
  });
  process.on('unhandledRejection', (reason) => {
    console.error('unhandledRejection', reason);
  });

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`HTML-to-PPT API listening on port ${PORT}`);
  });
}

module.exports = app;
