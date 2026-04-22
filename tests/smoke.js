const assert = require('assert');
const zlib = require('zlib');
const Module = require('module');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

function listZipEntries(buffer) {
  const entries = [];
  let offset = 0;

  while (offset <= buffer.length - 46) {
    const signature = buffer.readUInt32LE(offset);
    if (signature !== 0x02014b50) {
      offset += 1;
      continue;
    }

    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);

    const nameStart = offset + 46;
    const nameEnd = nameStart + fileNameLength;
    const fileName = buffer.toString('utf8', nameStart, nameEnd);

    const localSig = buffer.readUInt32LE(localHeaderOffset);
    if (localSig !== 0x04034b50) {
      throw new Error(`Invalid local header for ${fileName}`);
    }

    const localNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const dataEnd = dataStart + compressedSize;
    const compressed = buffer.subarray(dataStart, dataEnd);

    let data;
    if (compressionMethod === 0) {
      data = compressed;
    } else if (compressionMethod === 8) {
      data = zlib.inflateRawSync(compressed);
    } else {
      throw new Error(`Unsupported compression method ${compressionMethod} for ${fileName}`);
    }

    entries.push({ name: fileName, data });

    offset = nameEnd + extraLength + commentLength;
  }

  return entries;
}

function getEntryText(entries, name) {
  const entry = entries.find((item) => item.name === name);
  if (!entry) return null;
  return entry.data.toString('utf8');
}

async function createSamplePdfBuffer() {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  const first = pdfDoc.addPage([612, 792]);
  first.drawText('Quarterly update from PDF', {
    x: 72,
    y: 700,
    size: 24,
    font,
    color: rgb(0.1, 0.2, 0.4),
  });
  first.drawText('Revenue grew 17% year-over-year.', {
    x: 72,
    y: 660,
    size: 14,
    font,
    color: rgb(0.15, 0.15, 0.15),
  });

  const second = pdfDoc.addPage([612, 792]);
  second.drawRectangle({
    x: 72,
    y: 420,
    width: 468,
    height: 180,
    borderColor: rgb(0.2, 0.4, 0.7),
    borderWidth: 2,
    color: rgb(0.94, 0.96, 1),
  });

  return Buffer.from(await pdfDoc.save());
}

async function testSimpleModeNativeCss() {
  const { htmlToPptBuffer } = require('../converter');

  const html = `
    <style>
      .slide { width: 1280px; height: 720px; }
      h1 { color: rgb(220, 38, 38); text-align: center; margin: 40px 0 16px; font-size: 56px; }
      p { color: #1f2937; font-size: 24px; font-style: italic; line-height: 1.7; margin: 8px 0; }
      li { color: #0f766e; font-size: 22px; }
    </style>
    <div class="slide">
      <h1>Quarterly Update</h1>
      <p><strong>Revenue</strong> up 17% YoY.</p>
      <ul><li>New enterprise customers</li><li>Expanded product line</li></ul>
    </div>
    <div class="slide">
      <h2 style="font-size:42px;color:#111827">Second Slide</h2>
      <p>Body fallback check.</p>
    </div>
  `;

  const buffer = await htmlToPptBuffer(html, 'Simple Native', 'simple');
  assert(buffer.length > 0, 'Expected PPTX buffer for simple mode');

  const entries = listZipEntries(buffer);
  const slideFiles = entries.filter((entry) => entry.name.startsWith('ppt/slides/slide'));
  assert.strictEqual(slideFiles.length, 2, 'Expected two slides from .slide containers');

  const slide1 = getEntryText(entries, 'ppt/slides/slide1.xml');
  assert(slide1 && slide1.includes('Quarterly Update'), 'Slide 1 should include heading text');
  assert(slide1.includes('New enterprise customers'), 'Slide 1 should include list item text');
  assert(slide1.includes('<a:buChar'), 'List items should map to native bullets');
  assert(slide1.includes('srgbClr val="DC2626"'), 'Heading color should be mapped to PPT color');

  const slide2 = getEntryText(entries, 'ppt/slides/slide2.xml');
  assert(slide2 && slide2.includes('Second Slide'), 'Slide 2 should include second slide text');
}

async function testSimpleModeBodyFallback() {
  const { htmlToPptBuffer } = require('../converter');
  const buffer = await htmlToPptBuffer('<h1>Single Slide</h1><p>Uses body fallback.</p>', 'Body', 'simple');

  const entries = listZipEntries(buffer);
  const slideFiles = entries.filter((entry) => entry.name.startsWith('ppt/slides/slide'));
  assert.strictEqual(slideFiles.length, 1, 'Expected one slide when no .slide containers exist');

  const slide1 = getEntryText(entries, 'ppt/slides/slide1.xml');
  assert(slide1.includes('Single Slide'), 'Body fallback slide should include content');
}

async function testSimpleModeComplexCssLayout() {
  const { htmlToPptBuffer } = require('../converter');
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
    <style>
      html, body { margin:0; padding:0; width:100%; background:#fff; font-family: Inter, Arial, sans-serif; }
      .slide { width:1280px; height:720px; display:flex; flex-direction:column; background:#fff; }
      .slide-inner { flex:1; padding:24px; }
      .layout { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
      .roadmap, .funds { border:1px solid #D9E2F1; border-radius:16px; background:#F8FAFC; padding:16px; }
      .kicker { font-size:12px; letter-spacing:0.16em; text-transform:uppercase; color:#2563EB; font-weight:700; }
      .headline { font-size:28px; font-weight:800; color:#0F172A; margin:10px 0; }
      .copy { font-size:15px; line-height:1.45; color:#475569; }
      .box { padding:10px; border:1px solid #D9E2F1; border-radius:12px; background:#fff; margin-bottom:8px; }
      .slide-footer { height:44px; padding:0 24px 16px 24px; display:flex; justify-content:space-between; align-items:flex-end; font-size:16px; color:#0F172A; }
    </style>
    </head>
    <body>
      <section class="slide">
        <div class="slide-inner">
          <h1>Roadmap and use of funds</h1>
          <div class="layout">
            <div class="roadmap">
              <div class="kicker">Next 12 months</div>
              <h2 class="headline">Use funding to validate, then scale.</h2>
              <p class="copy">Validate product with customers and build a repeatable sales motion.</p>
            </div>
            <div class="funds">
              <div class="kicker">Use of funds</div>
              <div class="box"><strong>Product</strong><p>Model refinement and automation.</p></div>
              <div class="box"><strong>Go-to-market</strong><p>Sales materials and distribution.</p></div>
            </div>
          </div>
        </div>
        <div class="slide-footer"><span>AI Pilot 07</span><span>12</span></div>
      </section>
    </body>
    </html>
  `;

  const buffer = await htmlToPptBuffer(html, 'Complex CSS', 'simple');
  const entries = listZipEntries(buffer);
  const slide1 = getEntryText(entries, 'ppt/slides/slide1.xml');

  assert(slide1.includes('Roadmap and use of funds'), 'Should include heading text from complex layout');
  assert(slide1.includes('NEXT 12 MONTHS'), 'Should apply text-transform and include div text');
  assert(slide1.includes('AI Pilot 07'), 'Should include footer span text');
  assert(slide1.includes('<p:sp>'), 'Should include native shapes/textboxes for cards and boxes');
}

async function testSimpleModeInlineCssRuns() {
  const { htmlToPptBuffer } = require('../converter');
  const html = `
    <style>
      .slide { width: 1280px; height: 720px; padding: 48px; font-family: Arial, sans-serif; }
      p { font-size: 36px; color: #111827; }
      .accent { color: rgb(190, 24, 93); font-weight: 700; }
      .under { text-decoration: underline; }
      .strike { text-decoration: line-through; }
    </style>
    <section class="slide">
      <p><span class="accent">Revenue</span> <span class="under">up</span> <span class="strike">17%</span> YoY</p>
    </section>
  `;

  const buffer = await htmlToPptBuffer(html, 'Inline CSS', 'simple');
  const entries = listZipEntries(buffer);
  const slide1 = getEntryText(entries, 'ppt/slides/slide1.xml');

  assert(slide1.includes('Revenue'), 'Should retain inline span content in same text flow');
  assert(slide1.includes('srgbClr val="BE185D"'), 'Should map inline run color');
  assert(slide1.includes('u="sng"'), 'Should map underline text-decoration');
  assert(slide1.includes('strike="sngStrike"'), 'Should map line-through text-decoration');
}

async function testBrowserModeStillImageBacked() {
  const { htmlToPptBuffer } = require('../converter');
  const html = '<div class="slide" style="width:1280px;height:720px"><h1>Image slide</h1></div>';
  const buffer = await htmlToPptBuffer(html, 'Browser Mode', 'browser');

  const entries = listZipEntries(buffer);
  const slide1 = getEntryText(entries, 'ppt/slides/slide1.xml');
  assert(slide1.includes('<p:pic>'), 'Browser mode should still embed image-backed slides');
}

async function testApiHeadersUnchanged() {
  const app = require('../index');

  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));

  try {
    const { port } = server.address();

    const convertResponse = await fetch(`http://127.0.0.1:${port}/api/convert`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        title: 'API Test',
        filename: 'api-test',
        renderMode: 'simple',
        html: '<h1>API Check</h1><p>JSON endpoint</p>',
      }),
    });

    assert.strictEqual(convertResponse.status, 200, '/api/convert should return 200');
    assert.strictEqual(
      convertResponse.headers.get('content-type'),
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Type should remain PPTX MIME'
    );
    assert(
      (convertResponse.headers.get('content-disposition') || '').includes('api-test.pptx'),
      'Content-Disposition should include sanitized filename'
    );

    const rawResponse = await fetch(
      `http://127.0.0.1:${port}/api/convert-raw?title=Raw&filename=raw-test&renderMode=simple`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'text/html' },
        body: '<h1>Raw Endpoint</h1><p>Raw HTML endpoint</p>',
      }
    );

    assert.strictEqual(rawResponse.status, 200, '/api/convert-raw should return 200');
    assert(
      (rawResponse.headers.get('content-disposition') || '').includes('raw-test.pptx'),
      'Raw endpoint should preserve filename behavior'
    );
  } finally {
    await new Promise((resolve, reject) => {
      server.close((error) => (error ? reject(error) : resolve()));
    });
  }
}

async function testMissingBinaryErrorPath() {
  const originalLoad = Module._load;

  Module._load = function patchedLoad(request, parent, isMain) {
    if (request === 'playwright') {
      return {
        chromium: {
          launch: async () => {
            throw new Error("Executable doesn't exist at /fake/chromium");
          },
        },
      };
    }
    return originalLoad(request, parent, isMain);
  };

  const converterPath = require.resolve('../converter');
  delete require.cache[converterPath];

  try {
    const { htmlToPptBuffer } = require('../converter');

    await assert.rejects(
      () => htmlToPptBuffer('<p>missing browser</p>', 'Missing', 'simple'),
      /Simple render mode requires Playwright browser binaries/
    );
  } finally {
    Module._load = originalLoad;
    delete require.cache[converterPath];
  }
}

async function testPdfInternalValidationHelpers() {
  const { _internal } = require('../pdf-converter');

  assert.deepStrictEqual(_internal.parsePageRangeSpec('', 4), [1, 2, 3, 4], 'Empty range should select all pages');
  assert.deepStrictEqual(_internal.parsePageRangeSpec('1-2,4', 5), [1, 2, 4], 'Range parser should support mixed tokens');

  assert.throws(() => _internal.parsePageRangeSpec('2-1', 5), /Invalid pageRange/, 'Descending range should fail');
  assert.throws(() => _internal.parsePageRangeSpec('1-9', 3), /document bounds/, 'Out-of-bounds range should fail');

  const high = _internal.buildPageFidelity({
    textItems: [
      { text: 'Strong text signal', xPx: 10, yPx: 10, wPx: 100, hPx: 20 },
      { text: 'More text signal', xPx: 10, yPx: 40, wPx: 100, hPx: 20 },
    ],
  });
  assert.strictEqual(high.passed, true, 'Fidelity should pass with readable text + valid layout');

  const low = _internal.buildPageFidelity({
    textItems: [{ text: 'x', xPx: -999, yPx: -999, wPx: 0, hPx: 0 }],
  });
  assert.strictEqual(low.passed, false, 'Fidelity should fail with invalid geometry');
}

async function testPdfEndpointValidationAndHeaders() {
  const app = require('../index');
  const samplePdf = await createSamplePdfBuffer();

  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));

  try {
    const { port } = server.address();

    const noFileForm = new FormData();
    noFileForm.append('title', 'Missing File');
    const missingFileRes = await fetch(`http://127.0.0.1:${port}/api/convert-pdf`, {
      method: 'POST',
      body: noFileForm,
    });
    assert.strictEqual(missingFileRes.status, 400, 'Missing file should return 400');

    const wrongTypeForm = new FormData();
    wrongTypeForm.append('file', new Blob(['not a pdf'], { type: 'text/plain' }), 'bad.txt');
    const wrongTypeRes = await fetch(`http://127.0.0.1:${port}/api/convert-pdf`, {
      method: 'POST',
      body: wrongTypeForm,
    });
    assert.strictEqual(wrongTypeRes.status, 400, 'Non-PDF upload should return 400');

    const badRangeForm = new FormData();
    badRangeForm.append('file', new Blob([samplePdf], { type: 'application/pdf' }), 'sample.pdf');
    badRangeForm.append('pageRange', '4-2');
    const badRangeRes = await fetch(`http://127.0.0.1:${port}/api/convert-pdf`, {
      method: 'POST',
      body: badRangeForm,
    });
    assert.strictEqual(badRangeRes.status, 400, 'Invalid page range should return 400');
    const badRangeJson = await badRangeRes.json();
    assert(/Invalid pageRange/.test(badRangeJson.detail), 'Range error detail should be returned');

    const successForm = new FormData();
    successForm.append('title', 'PDF Upload');
    successForm.append('filename', 'pdf-upload');
    successForm.append('file', new Blob([samplePdf], { type: 'application/pdf' }), 'sample.pdf');
    const successRes = await fetch(`http://127.0.0.1:${port}/api/convert-pdf`, {
      method: 'POST',
      body: successForm,
    });

    assert.strictEqual(successRes.status, 200, 'PDF endpoint should return 200 for valid PDF');
    assert.strictEqual(
      successRes.headers.get('content-type'),
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'PDF endpoint should return PPTX MIME type'
    );
    assert(
      (successRes.headers.get('content-disposition') || '').includes('pdf-upload.pptx'),
      'PDF endpoint should include sanitized filename'
    );
    assert(successRes.headers.get('x-editable-pages'), 'Should return editable page count header');
    assert(successRes.headers.get('x-fallback-pages'), 'Should return fallback page count header');

    const pptBuffer = Buffer.from(await successRes.arrayBuffer());
    const entries = listZipEntries(pptBuffer);
    const slideFiles = entries.filter((entry) => entry.name.startsWith('ppt/slides/slide'));
    assert.strictEqual(slideFiles.length, 2, 'Sample 2-page PDF should produce 2 slides');

    const slide1 = getEntryText(entries, 'ppt/slides/slide1.xml');
    assert(slide1.includes('Quarterly update from PDF'), 'Editable text should appear in slide XML');

    const fallbackCount = Number(successRes.headers.get('x-fallback-pages'));
    assert(fallbackCount >= 1, 'Mixed sample should include at least one fallback image page');
  } finally {
    await new Promise((resolve, reject) => {
      server.close((error) => (error ? reject(error) : resolve()));
    });
  }
}

async function testPdfEndpointBrowserModeImageBacked() {
  const app = require('../index');
  const samplePdf = await createSamplePdfBuffer();

  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));

  try {
    const { port } = server.address();

    const form = new FormData();
    form.append('title', 'PDF Browser');
    form.append('filename', 'pdf-browser');
    form.append('renderMode', 'browser');
    form.append('file', new Blob([samplePdf], { type: 'application/pdf' }), 'sample.pdf');

    const res = await fetch(`http://127.0.0.1:${port}/api/convert-pdf`, {
      method: 'POST',
      body: form,
    });

    assert.strictEqual(res.status, 200, 'PDF browser mode should return 200');
    assert.strictEqual(res.headers.get('x-pdf-render-mode'), 'browser', 'Mode header should match request');
    assert.strictEqual(Number(res.headers.get('x-editable-pages')), 0, 'Browser mode should not produce editable pages');

    const pptBuffer = Buffer.from(await res.arrayBuffer());
    const entries = listZipEntries(pptBuffer);
    const slide1 = getEntryText(entries, 'ppt/slides/slide1.xml');
    assert(slide1.includes('<p:pic>'), 'Browser mode PDF slides should be image-backed');
  } finally {
    await new Promise((resolve, reject) => {
      server.close((error) => (error ? reject(error) : resolve()));
    });
  }
}

async function run() {
  const tests = [
    ['simple native css', testSimpleModeNativeCss],
    ['simple body fallback', testSimpleModeBodyFallback],
    ['simple complex css layout', testSimpleModeComplexCssLayout],
    ['simple inline css runs', testSimpleModeInlineCssRuns],
    ['browser image regression', testBrowserModeStillImageBacked],
    ['api headers unchanged', testApiHeadersUnchanged],
    ['pdf internal validation helpers', testPdfInternalValidationHelpers],
    ['pdf endpoint validation and headers', testPdfEndpointValidationAndHeaders],
    ['pdf endpoint browser mode image-backed', testPdfEndpointBrowserModeImageBacked],
    ['missing binary error path', testMissingBinaryErrorPath],
  ];

  for (const [name, fn] of tests) {
    await fn();
    process.stdout.write(`PASS ${name}\n`);
  }

  process.stdout.write('All smoke tests passed.\n');
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
