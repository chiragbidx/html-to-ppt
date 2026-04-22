const PptxGenJS = require('pptxgenjs');

const VIEWPORT = { width: 1280, height: 720 };
const SLIDE_SIZE_IN = { width: 13.333, height: 7.5 };

function pxToInX(px) {
  return (px / VIEWPORT.width) * SLIDE_SIZE_IN.width;
}

function pxToInY(px) {
  return (px / VIEWPORT.height) * SLIDE_SIZE_IN.height;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function cleanFontFace(fontFamily) {
  if (!fontFamily || typeof fontFamily !== 'string') return 'Calibri';
  const first = fontFamily.split(',')[0].replace(/['"]/g, '').trim();
  return first || 'Calibri';
}

function parseCssColor(color) {
  if (!color || typeof color !== 'string') {
    return { hex: '1F2937', alpha: 1, valid: false };
  }

  const input = color.trim();

  const hexMatch = input.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (hexMatch) {
    const raw = hexMatch[1];
    const hex = raw.length === 3
      ? `${raw[0]}${raw[0]}${raw[1]}${raw[1]}${raw[2]}${raw[2]}`.toUpperCase()
      : raw.toUpperCase();
    return { hex, alpha: 1, valid: true };
  }

  const rgbMatch = input.match(/^rgba?\(([^)]+)\)$/i);
  if (!rgbMatch) {
    return { hex: '1F2937', alpha: 1, valid: false };
  }

  const parts = rgbMatch[1].split(',').map((part) => part.trim());
  if (parts.length < 3) return { hex: '1F2937', alpha: 1, valid: false };

  const r = Number(parts[0]);
  const g = Number(parts[1]);
  const b = Number(parts[2]);
  const a = parts.length >= 4 ? Number(parts[3]) : 1;

  if ([r, g, b].some((v) => Number.isNaN(v))) {
    return { hex: '1F2937', alpha: 1, valid: false };
  }

  const hex = [r, g, b]
    .map((v) => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();

  return {
    hex,
    alpha: Number.isNaN(a) ? 1 : clamp(a, 0, 1),
    valid: true,
  };
}

function toPptAlign(textAlign) {
  if (textAlign === 'center') return 'center';
  if (textAlign === 'right') return 'right';
  if (textAlign === 'justify') return 'justify';
  return 'left';
}

function toRunOptions(style) {
  const fontSizePx = style.fontSizePx || 16;
  const fontSizePt = clamp(fontSizePx * 0.75, 8, 72);
  const weight = String(style.fontWeight || '400');
  const isBold = weight === 'bold' || Number(weight) >= 600;
  const color = parseCssColor(style.color);

  const options = {
    color: color.hex,
    bold: isBold,
    italic: String(style.fontStyle || '').includes('italic'),
    fontFace: cleanFontFace(style.fontFamily),
    fontSize: Number(fontSizePt.toFixed(2)),
  };

  const textDecorationLine = String(style.textDecorationLine || '');
  if (textDecorationLine.includes('underline')) {
    options.underline = { color: color.hex, style: 'sng' };
  }
  if (textDecorationLine.includes('line-through')) {
    options.strike = 'sngStrike';
  }

  const letterSpacingPx = Number(style.letterSpacingPx || 0);
  if (Number.isFinite(letterSpacingPx) && letterSpacingPx !== 0) {
    options.charSpace = Number((letterSpacingPx * 0.75).toFixed(2));
  }

  return options;
}

function createPpt(title) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'html-to-ppt API';
  pptx.subject = 'HTML conversion';
  pptx.title = title || 'Generated Presentation';
  return pptx;
}

function addNoContentSlide(pptx, message) {
  const slide = pptx.addSlide();
  slide.addText(message, {
    x: 0.7,
    y: 3.0,
    w: 12,
    h: 0.6,
    fontFace: 'Calibri',
    fontSize: 18,
    color: '6B7280',
    italic: true,
    align: 'center',
  });
}

function normalizePlaywrightMissingBinaryError(error, renderMode) {
  const message = error && error.message ? error.message : String(error);
  if (message.includes("Executable doesn't exist")) {
    throw new Error(
      `${renderMode} render mode requires Playwright browser binaries. Run: npx playwright install chromium`
    );
  }
  throw error;
}

function addDecorations(slide, decorations) {
  decorations.forEach((box) => {
    const fillColor = parseCssColor(box.backgroundColor);
    const borderColor = parseCssColor(box.borderColor);
    const boxOpacity = clamp(Number(box.opacity || 1), 0, 1);

    const borderWidthPt = Number(((box.borderWidthPx || 0) * 0.75).toFixed(2));
    const hasFill = fillColor.valid && fillColor.alpha > 0 && boxOpacity > 0;
    const hasBorder = borderColor.valid && borderColor.alpha > 0 && borderWidthPt > 0 && boxOpacity > 0;

    if (!hasFill && !hasBorder) return;

    const shapeType = PptxGenJS.ShapeType || {};
    const shapeName = box.radiusPx > 1 ? shapeType.roundRect || 'roundRect' : shapeType.rect || 'rect';
    const shapeOptions = {
      x: Number(pxToInX(box.xPx).toFixed(3)),
      y: Number(pxToInY(box.yPx).toFixed(3)),
      w: Number(pxToInX(box.wPx).toFixed(3)),
      h: Number(pxToInY(box.hPx).toFixed(3)),
    };

    if (hasFill) {
      shapeOptions.fill = {
        color: fillColor.hex,
        transparency: Number(((1 - (fillColor.alpha * boxOpacity)) * 100).toFixed(2)),
      };
    } else {
      shapeOptions.fill = { color: 'FFFFFF', transparency: 100 };
    }

    if (hasBorder) {
      shapeOptions.line = {
        color: borderColor.hex,
        pt: clamp(borderWidthPt, 0.25, 12),
        transparency: Number(((1 - (borderColor.alpha * boxOpacity)) * 100).toFixed(2)),
      };
    } else {
      shapeOptions.line = { color: 'FFFFFF', transparency: 100, pt: 0 };
    }

    slide.addShape(shapeName, shapeOptions);
  });
}

function addTexts(slide, textItems) {
  textItems.forEach((item) => {
    const runs = item.runs
      .filter((run) => run.text && run.text.trim().length > 0)
      .map((run) => ({ text: run.text, options: toRunOptions(run.style) }));

    if (!runs.length) return;

    const baseFont = item.runs[0] && item.runs[0].style ? item.runs[0].style.fontSizePx || 16 : 16;
    const lineHeightPx = Number(item.lineHeightPx || baseFont * 1.2);
    const lineSpacingMultiple = clamp(lineHeightPx / Math.max(baseFont, 1), 0.8, 3.0);

    const textOptions = {
      x: Number(pxToInX(item.xPx).toFixed(3)),
      y: Number(pxToInY(item.yPx).toFixed(3)),
      w: Number(pxToInX(item.wPx).toFixed(3)),
      h: Number(pxToInY(Math.max(item.hPx, lineHeightPx)).toFixed(3)),
      align: toPptAlign(item.textAlign),
      valign: 'top',
      breakLine: false,
      fit: 'shrink',
      lineSpacingMultiple: Number(lineSpacingMultiple.toFixed(2)),
    };

    if (item.isListItem) {
      textOptions.bullet = { indent: 18 };
    }

    slide.addText(runs, textOptions);
  });
}

async function extractNativeSlideModel(page) {
  return page.evaluate(() => {
    const SKIP_TAGS = new Set(['script', 'style', 'noscript', 'head', 'meta', 'link', 'title']);

    function toPx(value, fallback) {
      const num = Number.parseFloat(value || '');
      return Number.isFinite(num) ? num : fallback;
    }

    function normalizeWhitespace(text) {
      return (text || '').replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
    }

    function hasVisibleBox(rect, css) {
      if (rect.width <= 0 || rect.height <= 0) return false;
      if (css.display === 'none') return false;
      if (css.visibility === 'hidden') return false;
      const opacity = Number.parseFloat(css.opacity || '1');
      if (Number.isFinite(opacity) && opacity <= 0) return false;
      return true;
    }

    function applyTextTransform(text, textTransform) {
      if (textTransform === 'uppercase') return text.toUpperCase();
      if (textTransform === 'lowercase') return text.toLowerCase();
      if (textTransform === 'capitalize') {
        return text.replace(/\b([a-z])/g, (m) => m.toUpperCase());
      }
      return text;
    }

    function collectDecorations(container, containerRect) {
      const decorations = [];
      const elements = [container, ...Array.from(container.querySelectorAll('*'))];

      for (const el of elements) {
        const tag = el.tagName ? el.tagName.toLowerCase() : '';
        if (!tag || SKIP_TAGS.has(tag)) continue;

        const css = window.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        if (!hasVisibleBox(rect, css)) continue;

        const backgroundColor = css.backgroundColor;
        const borderColor = css.borderColor;
        const borderWidthPx = Math.max(
          toPx(css.borderTopWidth, 0),
          toPx(css.borderRightWidth, 0),
          toPx(css.borderBottomWidth, 0),
          toPx(css.borderLeftWidth, 0)
        );
        const radiusPx = Math.max(
          toPx(css.borderTopLeftRadius, 0),
          toPx(css.borderTopRightRadius, 0),
          toPx(css.borderBottomLeftRadius, 0),
          toPx(css.borderBottomRightRadius, 0)
        );

        const hasFill = backgroundColor && !backgroundColor.includes('rgba(0, 0, 0, 0)') && backgroundColor !== 'transparent';
        const hasBorder = borderWidthPx > 0 && borderColor && !borderColor.includes('rgba(0, 0, 0, 0)') && borderColor !== 'transparent';

        if (!hasFill && !hasBorder) continue;

        decorations.push({
          xPx: rect.left - containerRect.left,
          yPx: rect.top - containerRect.top,
          wPx: rect.width,
          hPx: rect.height,
          backgroundColor,
          borderColor,
          borderWidthPx,
          radiusPx,
          opacity: Number.parseFloat(css.opacity || '1'),
        });
      }

      return decorations;
    }

    function isInlineDisplay(displayValue) {
      if (!displayValue) return false;
      return displayValue.startsWith('inline') || displayValue === 'contents';
    }

    function normalizeNodeText(text) {
      const input = String(text || '');
      if (!input) return '';

      const hasLeadingSpace = /^\s/.test(input);
      const hasTrailingSpace = /\s$/.test(input);
      const collapsed = input.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();

      if (!collapsed) return '';

      return `${hasLeadingSpace ? ' ' : ''}${collapsed}${hasTrailingSpace ? ' ' : ''}`;
    }

    function isNestedInsideAnotherBlock(node, rootContainer) {
      let current = node.parentElement;
      while (current && current !== rootContainer) {
        const css = window.getComputedStyle(current);
        if (!isInlineDisplay(css.display)) return true;
        current = current.parentElement;
      }
      return false;
    }

    function mergeAdjacentRuns(runs) {
      if (!runs.length) return runs;
      const merged = [runs[0]];

      for (let i = 1; i < runs.length; i += 1) {
        const prev = merged[merged.length - 1];
        const curr = runs[i];
        const sameStyle =
          prev.style.color === curr.style.color
          && prev.style.fontSizePx === curr.style.fontSizePx
          && prev.style.fontWeight === curr.style.fontWeight
          && prev.style.fontStyle === curr.style.fontStyle
          && prev.style.fontFamily === curr.style.fontFamily
          && prev.style.letterSpacingPx === curr.style.letterSpacingPx
          && prev.style.textDecorationLine === curr.style.textDecorationLine;

        if (sameStyle) {
          prev.text += curr.text;
        } else {
          merged.push(curr);
        }
      }

      return merged;
    }

    function collectTextItems(container, containerRect) {
      const textItems = [];
      const elements = Array.from(container.querySelectorAll('*'));

      for (const el of elements) {
        const tag = el.tagName ? el.tagName.toLowerCase() : '';
        if (!tag || SKIP_TAGS.has(tag)) continue;

        const css = window.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        if (!hasVisibleBox(rect, css)) continue;
        if (isInlineDisplay(css.display)) continue;

        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        const textNodes = [];
        while (walker.nextNode()) {
          const node = walker.currentNode;
          if (!node || !node.parentElement) continue;
          if (isNestedInsideAnotherBlock(node, el)) continue;
          if (normalizeWhitespace(node.textContent).length === 0) continue;
          textNodes.push(node);
        }

        if (!textNodes.length) continue;

        let minLeft = Number.POSITIVE_INFINITY;
        let minTop = Number.POSITIVE_INFINITY;
        let maxRight = Number.NEGATIVE_INFINITY;
        let maxBottom = Number.NEGATIVE_INFINITY;
        const runs = [];

        for (const node of textNodes) {
          const parentCss = window.getComputedStyle(node.parentElement);
          const nodeRange = document.createRange();
          nodeRange.selectNodeContents(node);
          const nodeRect = nodeRange.getBoundingClientRect();
          if (nodeRect.width <= 0 || nodeRect.height <= 0) continue;

          minLeft = Math.min(minLeft, nodeRect.left);
          minTop = Math.min(minTop, nodeRect.top);
          maxRight = Math.max(maxRight, nodeRect.right);
          maxBottom = Math.max(maxBottom, nodeRect.bottom);

          const normalized = normalizeNodeText(node.textContent || '');
          if (!normalized) continue;
          const transformedText = applyTextTransform(normalized, parentCss.textTransform);
          if (!transformedText) continue;

          runs.push({
            text: transformedText,
            style: {
              color: parentCss.color,
              fontSizePx: toPx(parentCss.fontSize, toPx(css.fontSize, 16)),
              fontWeight: parentCss.fontWeight,
              fontStyle: parentCss.fontStyle,
              fontFamily: parentCss.fontFamily,
              letterSpacingPx: toPx(parentCss.letterSpacing, 0),
              textDecorationLine: parentCss.textDecorationLine,
            },
          });
        }

        if (!runs.length) continue;
        if (![minLeft, minTop, maxRight, maxBottom].every(Number.isFinite)) continue;

        const xPx = minLeft - containerRect.left;
        const yPx = minTop - containerRect.top;
        const wPx = maxRight - minLeft;
        const hPx = maxBottom - minTop;

        if (wPx <= 0 || hPx <= 0) continue;

        textItems.push({
          xPx,
          yPx,
          wPx,
          hPx,
          textAlign: css.textAlign,
          lineHeightPx: toPx(css.lineHeight, toPx(css.fontSize, 16) * 1.2),
          isListItem: tag === 'li',
          runs: mergeAdjacentRuns(runs),
        });
      }

      return textItems;
    }

    const slideRoots = Array.from(document.querySelectorAll('.slide'));
    const containers = slideRoots.length > 0 ? slideRoots : [document.body];

    return containers.map((container) => {
      const containerRect = container.getBoundingClientRect();
      const decorations = collectDecorations(container, containerRect);
      const textItems = collectTextItems(container, containerRect);

      return {
        decorations,
        textItems,
      };
    });
  });
}

async function htmlToPptBufferSimple(html, title) {
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ headless: true });

  try {
    const page = await browser.newPage({ viewport: VIEWPORT });
    await page.setContent(html || '', { waitUntil: 'networkidle' });
    await page.evaluate(() => document.fonts && document.fonts.ready);

    const slideModels = await extractNativeSlideModel(page);
    const hasContent = slideModels.some((slideModel) => {
      return slideModel.textItems.length > 0 || slideModel.decorations.length > 0;
    });

    const pptx = createPpt(title);

    if (!hasContent) {
      addNoContentSlide(pptx, 'No visible content extracted from provided HTML.');
      return pptx.write({ outputType: 'nodebuffer' });
    }

    slideModels.forEach((slideModel) => {
      if (!slideModel.textItems.length && !slideModel.decorations.length) return;

      const slide = pptx.addSlide();
      addDecorations(slide, slideModel.decorations);
      addTexts(slide, slideModel.textItems);
    });

    return pptx.write({ outputType: 'nodebuffer' });
  } finally {
    await browser.close();
  }
}

async function htmlToPptBufferWithBrowser(html, title) {
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ headless: true });

  try {
    const page = await browser.newPage({ viewport: { width: 1280, height: 720 } });
    await page.setContent(html || '', { waitUntil: 'networkidle' });
    await page.evaluate(() => document.fonts && document.fonts.ready);

    const slideCount = await page.$$eval('.slide', (nodes) => nodes.length);
    const screenshots = [];

    if (slideCount > 0) {
      const elements = await page.$$('.slide');
      for (const el of elements) {
        const shot = await el.screenshot({ type: 'png' });
        screenshots.push(shot);
      }
    } else {
      const body = await page.$('body');
      if (body) {
        const shot = await body.screenshot({ type: 'png' });
        screenshots.push(shot);
      }
    }

    const pptx = createPpt(title);

    if (!screenshots.length) {
      addNoContentSlide(pptx, 'No renderable content found in HTML.');
    } else {
      screenshots.forEach((imageBuffer) => {
        const slide = pptx.addSlide();
        slide.addImage({
          data: `image/png;base64,${imageBuffer.toString('base64')}`,
          x: 0,
          y: 0,
          w: 13.333,
          h: 7.5,
        });
      });
    }

    return pptx.write({ outputType: 'nodebuffer' });
  } finally {
    await browser.close();
  }
}

async function htmlToPptBufferAuto(html, title, renderMode = 'browser') {
  if (renderMode === 'simple') {
    try {
      return await htmlToPptBufferSimple(html, title);
    } catch (error) {
      normalizePlaywrightMissingBinaryError(error, 'Simple');
    }
  }

  if (renderMode === 'browser') {
    try {
      return await htmlToPptBufferWithBrowser(html, title);
    } catch (error) {
      normalizePlaywrightMissingBinaryError(error, 'Browser');
    }
  }

  throw new Error('Invalid renderMode. Use "browser" or "simple".');
}

module.exports = {
  htmlToPptBuffer: htmlToPptBufferAuto,
  _internal: {
    normalizePlaywrightMissingBinaryError,
  },
};
