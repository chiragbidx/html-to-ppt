const { createCanvas } = require('@napi-rs/canvas');
const PptxGenJS = require('pptxgenjs');

const VIEWPORT = { width: 1280, height: 720 };
const SLIDE_SIZE_IN = { width: 13.333, height: 7.5 };
const MAX_PDF_PAGES = 100;
const MAX_PDF_SIZE_BYTES = 15 * 1024 * 1024;
const DEFAULT_FONT_FACE = 'Calibri';

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function normalizeWhitespace(text) {
  return String(text || '').replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
}

function pxToInX(px) {
  return (px / VIEWPORT.width) * SLIDE_SIZE_IN.width;
}

function pxToInY(px) {
  return (px / VIEWPORT.height) * SLIDE_SIZE_IN.height;
}

function pointsToPx(points) {
  return points * (96 / 72);
}

function cleanFontFace(fontFamily) {
  if (!fontFamily || typeof fontFamily !== 'string') return DEFAULT_FONT_FACE;
  const first = fontFamily.split(',')[0].replace(/['"]/g, '').trim();
  return first || DEFAULT_FONT_FACE;
}

function rgbToHex(rgb) {
  if (!Array.isArray(rgb) || rgb.length < 3) return '1F2937';
  return rgb
    .slice(0, 3)
    .map((v) => clamp(Math.round(Number(v) * 255), 0, 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

function createPpt(title) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'html-to-ppt API';
  pptx.subject = 'PDF conversion';
  pptx.title = title || 'Generated Presentation';
  return pptx;
}

function parsePageRangeSpec(pageRangeSpec, totalPages) {
  if (!pageRangeSpec || String(pageRangeSpec).trim() === '') {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }

  const selected = new Set();
  const chunks = String(pageRangeSpec)
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean);

  if (!chunks.length) {
    throw new Error('Invalid pageRange. Use format like "1-5,8".');
  }

  for (const chunk of chunks) {
    const rangeMatch = chunk.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const start = Number(rangeMatch[1]);
      const end = Number(rangeMatch[2]);
      if (!Number.isInteger(start) || !Number.isInteger(end) || start < 1 || end < 1 || start > end) {
        throw new Error('Invalid pageRange. Use positive ascending ranges like "1-5".');
      }
      for (let i = start; i <= end; i += 1) {
        selected.add(i);
      }
      continue;
    }

    if (!/^\d+$/.test(chunk)) {
      throw new Error('Invalid pageRange. Use format like "1-5,8".');
    }

    selected.add(Number(chunk));
  }

  const pages = Array.from(selected).sort((a, b) => a - b);
  if (!pages.length) throw new Error('pageRange selected no pages.');
  if (pages[0] < 1 || pages[pages.length - 1] > totalPages) {
    throw new Error(`pageRange exceeds document bounds (1-${totalPages}).`);
  }

  return pages;
}

function isLikelyBold(fontName = '') {
  return /bold|black|heavy|demi|semibold/i.test(fontName);
}

function isLikelyItalic(fontName = '') {
  return /italic|oblique/i.test(fontName);
}

function mapPdfTextItemToPpt(item, viewport) {
  const [a, b, c, d, e, f] = item.transform || [1, 0, 0, 1, 0, 0];
  const fontScale = Math.max(Math.hypot(a, b), Math.hypot(c, d), 1);
  const rawHeightPt = Number(item.height) > 0 ? Number(item.height) : fontScale;
  const rawWidthPt = Number(item.width) > 0 ? Number(item.width) : Math.max((item.str || '').length * rawHeightPt * 0.45, rawHeightPt);

  const xPt = Number(e) || 0;
  const yPt = Number(f) || 0;

  const xPx = pointsToPx(xPt);
  const baselineYPx = pointsToPx(viewport.height - yPt);
  const hPx = Math.max(pointsToPx(rawHeightPt), 8);
  const yPx = baselineYPx - hPx;
  const wPx = Math.max(pointsToPx(rawWidthPt), 8);

  const str = normalizeWhitespace(item.str || '');
  if (!str) return null;

  return {
    text: str,
    xPx,
    yPx,
    wPx,
    hPx,
    fontSizePt: clamp(rawHeightPt, 6, 72),
  };
}

function buildPageFidelity(pageModel) {
  const totalChars = pageModel.textItems.reduce((sum, item) => sum + item.text.length, 0);
  const inBoundsCount = pageModel.textItems.filter((item) => {
    return item.xPx >= -50 && item.yPx >= -50 && item.wPx > 0 && item.hPx > 0;
  }).length;

  const positionedRatio = pageModel.textItems.length === 0 ? 0 : inBoundsCount / pageModel.textItems.length;
  const charScore = clamp(totalChars / 80, 0, 1);
  const layoutScore = clamp(positionedRatio, 0, 1);
  const score = Number((charScore * 0.65 + layoutScore * 0.35).toFixed(3));

  const passed = pageModel.textItems.length > 0 && layoutScore >= 0.85 && score >= 0.45;

  return {
    score,
    totalChars,
    layoutScore,
    passed,
  };
}

function addTextItems(slide, textItems) {
  textItems.forEach((item) => {
    const options = {
      x: Number(pxToInX(item.xPx).toFixed(3)),
      y: Number(pxToInY(item.yPx).toFixed(3)),
      w: Number(pxToInX(item.wPx).toFixed(3)),
      h: Number(pxToInY(item.hPx).toFixed(3)),
      fontFace: cleanFontFace(item.fontFace),
      fontSize: Number(item.fontSizePt.toFixed(2)),
      color: item.colorHex || '1F2937',
      bold: Boolean(item.bold),
      italic: Boolean(item.italic),
      breakLine: false,
      valign: 'top',
      fit: 'shrink',
      margin: 0,
    };

    slide.addText(item.text, options);
  });
}

function addVectorItems(slide, vectorItems) {
  const shapeType = PptxGenJS.ShapeType || {};
  const lineType = shapeType.line || 'line';
  const rectType = shapeType.rect || 'rect';

  vectorItems.forEach((shape) => {
    if (shape.kind === 'line') {
      slide.addShape(lineType, {
        x: Number(pxToInX(shape.xPx).toFixed(3)),
        y: Number(pxToInY(shape.yPx).toFixed(3)),
        w: Number(pxToInX(shape.wPx).toFixed(3)),
        h: Number(pxToInY(shape.hPx).toFixed(3)),
        line: {
          color: shape.strokeHex || '4B5563',
          pt: Number((shape.lineWidthPt || 0.75).toFixed(2)),
        },
      });
      return;
    }

    if (shape.kind === 'rect') {
      slide.addShape(rectType, {
        x: Number(pxToInX(shape.xPx).toFixed(3)),
        y: Number(pxToInY(shape.yPx).toFixed(3)),
        w: Number(pxToInX(shape.wPx).toFixed(3)),
        h: Number(pxToInY(shape.hPx).toFixed(3)),
        fill: shape.fillHex ? { color: shape.fillHex } : { color: 'FFFFFF', transparency: 100 },
        line: shape.strokeHex
          ? {
              color: shape.strokeHex,
              pt: Number((shape.lineWidthPt || 0.75).toFixed(2)),
            }
          : { color: 'FFFFFF', transparency: 100, pt: 0 },
      });
    }
  });
}

async function renderPdfPageImage(pdfjsLib, doc, pageNumber, targetWidthPx = VIEWPORT.width) {
  const onePixelWhitePng = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9X6Ys6kAAAAASUVORK5CYII=',
    'base64'
  );

  try {
    const page = await doc.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1.0 });
    const safeWidth = Math.max(1, baseViewport.width);
    const scale = clamp(targetWidthPx / safeWidth, 0.5, 3.0);
    const viewport = page.getViewport({ scale });

    const canvas = createCanvas(Math.max(1, Math.ceil(viewport.width)), Math.max(1, Math.ceil(viewport.height)));
    const context = canvas.getContext('2d');

    context.fillStyle = '#FFFFFF';
    context.fillRect(0, 0, canvas.width, canvas.height);

    const canvasFactory = {
      create(width, height) {
        const pageCanvas = createCanvas(Math.max(1, Math.ceil(width)), Math.max(1, Math.ceil(height)));
        return { canvas: pageCanvas, context: pageCanvas.getContext('2d') };
      },
      reset(target, width, height) {
        target.canvas.width = Math.max(1, Math.ceil(width));
        target.canvas.height = Math.max(1, Math.ceil(height));
      },
      destroy(target) {
        target.canvas.width = 0;
        target.canvas.height = 0;
        target.canvas = null;
        target.context = null;
      },
    };

    await page.render({
      canvasContext: context,
      viewport,
      canvasFactory,
    }).promise;

    return canvas.toBuffer('image/png');
  } catch (_error) {
    return onePixelWhitePng;
  }
}

async function loadPdfJs() {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  return pdfjs;
}

async function extractPageModel(pdfjsLib, doc, pageNumber) {
  const page = await doc.getPage(pageNumber);
  const viewport = page.getViewport({ scale: 1.0 });

  const textContent = await page.getTextContent();
  const textItems = [];

  for (const item of textContent.items || []) {
    const mapped = mapPdfTextItemToPpt(item, viewport);
    if (!mapped) continue;

    const style = textContent.styles && item.fontName ? textContent.styles[item.fontName] : null;
    mapped.fontFace = cleanFontFace((style && style.fontFamily) || item.fontName || DEFAULT_FONT_FACE);
    mapped.colorHex = '1F2937';
    mapped.bold = isLikelyBold(item.fontName);
    mapped.italic = isLikelyItalic(item.fontName);
    textItems.push(mapped);
  }

  const vectorItems = [];
  try {
    const operatorList = await page.getOperatorList();
    const OPS = pdfjsLib.OPS || {};

    let strokeHex = '4B5563';
    let fillHex = null;
    let lineWidthPt = 0.75;
    let pendingRects = [];

    const flushRects = (fill, stroke) => {
      for (const rect of pendingRects) {
        vectorItems.push({
          kind: 'rect',
          xPx: pointsToPx(rect.x),
          yPx: pointsToPx(viewport.height - rect.y - rect.h),
          wPx: pointsToPx(rect.w),
          hPx: pointsToPx(rect.h),
          fillHex: fill,
          strokeHex: stroke,
          lineWidthPt,
        });
      }
      pendingRects = [];
    };

    for (let i = 0; i < operatorList.fnArray.length; i += 1) {
      const fn = operatorList.fnArray[i];
      const args = operatorList.argsArray[i];

      if (fn === OPS.setStrokeRGBColor) {
        strokeHex = rgbToHex(args);
      } else if (fn === OPS.setFillRGBColor) {
        fillHex = rgbToHex(args);
      } else if (fn === OPS.setLineWidth) {
        lineWidthPt = clamp(Number(args && args[0]) || 0.75, 0.25, 12);
      } else if (fn === OPS.constructPath) {
        const [ops, coords] = args || [];
        if (!Array.isArray(ops) || !Array.isArray(coords)) continue;
        let cursor = 0;
        for (const opCode of ops) {
          if (opCode === OPS.rectangle) {
            const x = Number(coords[cursor]);
            const y = Number(coords[cursor + 1]);
            const w = Number(coords[cursor + 2]);
            const h = Number(coords[cursor + 3]);
            cursor += 4;
            if ([x, y, w, h].every(Number.isFinite) && Math.abs(w) > 0.5 && Math.abs(h) > 0.5) {
              pendingRects.push({ x, y, w: Math.abs(w), h: Math.abs(h) });
            }
          } else if (opCode === OPS.moveTo || opCode === OPS.lineTo) {
            cursor += 2;
          } else if (opCode === OPS.curveTo || opCode === OPS.curveTo2 || opCode === OPS.curveTo3) {
            cursor += 6;
          } else if (opCode === OPS.closePath) {
            // no-op
          }
        }
      } else if (fn === OPS.fill || fn === OPS.eoFill) {
        flushRects(fillHex, null);
      } else if (fn === OPS.stroke) {
        flushRects(null, strokeHex);
      } else if (fn === OPS.fillStroke || fn === OPS.eoFillStroke) {
        flushRects(fillHex, strokeHex);
      } else if (fn === OPS.closeFillStroke || fn === OPS.closeEOFillStroke) {
        flushRects(fillHex, strokeHex);
      }
    }

    if (pendingRects.length) {
      flushRects(fillHex, strokeHex);
    }
  } catch (_error) {
    // Vector extraction is best-effort in phase 1.
  }

  const dedupedText = textItems.filter((item, idx) => {
    const prev = textItems[idx - 1];
    if (!prev) return true;
    return !(
      item.text === prev.text
      && Math.abs(item.xPx - prev.xPx) < 0.5
      && Math.abs(item.yPx - prev.yPx) < 0.5
      && Math.abs(item.wPx - prev.wPx) < 0.5
      && Math.abs(item.hPx - prev.hPx) < 0.5
    );
  });

  return {
    pageNumber,
    textItems: dedupedText,
    vectorItems,
    fidelity: buildPageFidelity({ textItems: dedupedText }),
  };
}

async function pdfToPptBuffer(pdfBuffer, options = {}) {
  const { title, pageRange, renderMode = 'simple' } = options;
  if (!Buffer.isBuffer(pdfBuffer)) {
    throw new Error('PDF input must be a Buffer.');
  }

  if (pdfBuffer.length === 0) {
    throw new Error('PDF file is empty.');
  }

  if (pdfBuffer.length > MAX_PDF_SIZE_BYTES) {
    throw new Error(`PDF exceeds max size of ${Math.floor(MAX_PDF_SIZE_BYTES / (1024 * 1024))}MB.`);
  }

  const pdfjsLib = await loadPdfJs();
  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(pdfBuffer), useSystemFonts: true });
  const doc = await loadingTask.promise;

  if (doc.numPages > MAX_PDF_PAGES) {
    throw new Error(`PDF exceeds max pages (${MAX_PDF_PAGES}).`);
  }

  if (renderMode !== 'simple' && renderMode !== 'browser') {
    throw new Error('Invalid renderMode. Use "simple" or "browser".');
  }

  const pageNumbers = parsePageRangeSpec(pageRange, doc.numPages);
  const pptx = createPpt(title);

  let editablePages = 0;
  let fallbackPages = 0;

  for (const pageNumber of pageNumbers) {
    if (renderMode === 'browser') {
      const image = await renderPdfPageImage(pdfjsLib, doc, pageNumber);
      const slide = pptx.addSlide();
      slide.addImage({
        data: `image/png;base64,${image.toString('base64')}`,
        x: 0,
        y: 0,
        w: SLIDE_SIZE_IN.width,
        h: SLIDE_SIZE_IN.height,
      });
      fallbackPages += 1;
      continue;
    }

    const pageModel = await extractPageModel(pdfjsLib, doc, pageNumber);
    if (pageModel.fidelity.passed) {
      const slide = pptx.addSlide();
      addVectorItems(slide, pageModel.vectorItems);
      addTextItems(slide, pageModel.textItems);
      editablePages += 1;
      continue;
    }

    const image = await renderPdfPageImage(pdfjsLib, doc, pageNumber);
    const slide = pptx.addSlide();
    slide.addImage({
      data: `image/png;base64,${image.toString('base64')}`,
      x: 0,
      y: 0,
      w: SLIDE_SIZE_IN.width,
      h: SLIDE_SIZE_IN.height,
    });
    fallbackPages += 1;
  }

  const buffer = await pptx.write({ outputType: 'nodebuffer' });

  return {
    buffer,
    metadata: {
      editablePages,
      fallbackPages,
      totalSlides: pageNumbers.length,
      selectedPages: pageNumbers,
    },
  };
}

module.exports = {
  pdfToPptBuffer,
  MAX_PDF_SIZE_BYTES,
  MAX_PDF_PAGES,
  _internal: {
    parsePageRangeSpec,
    mapPdfTextItemToPpt,
    buildPageFidelity,
    rgbToHex,
  },
};
