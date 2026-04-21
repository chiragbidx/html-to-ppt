const PptxGenJS = require('pptxgenjs');

function cleanText(text) {
  return (text || '').replace(/\s+/g, ' ').trim();
}

function decodeHtmlEntities(text) {
  if (!text) return '';
  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

function stripTags(html) {
  return decodeHtmlEntities(
    (html || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]*>/g, ' ')
  );
}

function extractBlocks(html) {
  const blocks = [];
  const source = html || '';

  function pushBlock(kind, text, level = 0) {
    const value = cleanText(text);
    if (value) {
      blocks.push({ kind, text: value, level });
    }
  }

  const blockRegex = /<(h1|h2|h3|p|li)\b[^>]*>([\s\S]*?)<\/\1>/gi;
  let match;
  while ((match = blockRegex.exec(source)) !== null) {
    const tag = match[1].toLowerCase();
    const text = stripTags(match[2]);
    if (tag === 'li') {
      pushBlock('li', `\u2022 ${text}`);
    } else {
      pushBlock(tag, text);
    }
  }

  return blocks;
}

async function htmlToPptBufferSimple(html, title) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'html-to-ppt API';
  pptx.subject = 'HTML conversion';
  pptx.title = title || 'Generated Presentation';

  const slide = pptx.addSlide();
  const blocks = extractBlocks(html);

  let y = 0.4;
  const maxY = 6.8;

  for (const block of blocks) {
    let fontSize = 18;
    let bold = false;
    let color = '1F2937';
    let indent = 0;

    if (block.kind === 'h1') {
      fontSize = 32;
      bold = true;
      color = '111827';
    } else if (block.kind === 'h2') {
      fontSize = 26;
      bold = true;
    } else if (block.kind === 'h3') {
      fontSize = 22;
      bold = true;
    } else if (block.kind === 'li') {
      fontSize = 16;
      indent = 0.2;
    }

    const h = Math.max(0.28, Math.min(1.0, block.text.length / 120 + 0.25));
    if (y + h > maxY) break;

    slide.addText(block.text, {
      x: 0.5 + indent,
      y,
      w: 12.3 - indent,
      h,
      fontFace: 'Calibri',
      fontSize,
      bold,
      color,
      valign: 'top',
      breakLine: true,
    });

    y += h + 0.08;
  }

  if (!blocks.length) {
    slide.addText('No visible content extracted from provided HTML.', {
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

  return pptx.write({ outputType: 'nodebuffer' });
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

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'html-to-ppt API';
    pptx.subject = 'HTML conversion';
    pptx.title = title || 'Generated Presentation';

    if (!screenshots.length) {
      const slide = pptx.addSlide();
      slide.addText('No renderable content found in HTML.', {
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
    return htmlToPptBufferSimple(html, title);
  }

  if (renderMode === 'browser') {
    try {
      return await htmlToPptBufferWithBrowser(html, title);
    } catch (error) {
      const message = error && error.message ? error.message : String(error);
      if (message.includes('Executable doesn\'t exist')) {
        throw new Error(
          'Browser render mode requires Playwright browser binaries. Run: npx playwright install chromium'
        );
      }
      throw error;
    }
  }

  throw new Error('Invalid renderMode. Use "browser" or "simple".');
}

module.exports = { htmlToPptBuffer: htmlToPptBufferAuto };
