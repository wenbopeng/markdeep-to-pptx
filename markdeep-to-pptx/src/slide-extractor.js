/**
 * Slide Extractor - Extract slide content from rendered Markdeep Slides HTML
 * 
 * This module uses Playwright to open a Markdeep Slides HTML file,
 * wait for it to fully render, and extract the structured slide content
 * for conversion to PPTX.
 */

import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Extract slide content from a rendered Markdeep Slides HTML file
 * @param {string} htmlPath - Path to the HTML file
 * @returns {Promise<Object>} - Extracted slide data including slides array and metadata
 */
export async function extractSlides(htmlPath) {
    const browser = await chromium.launch({
        headless: true,
        args: ['--disable-web-security', '--allow-file-access-from-files']
    });

    const context = await browser.newContext({
        viewport: { width: 1920, height: 1080 }
    });

    const page = await context.newPage();

    // Convert to file:// URL if it's a local path
    let fileUrl = htmlPath;
    if (!htmlPath.startsWith('file://') && !htmlPath.startsWith('http')) {
        const absolutePath = path.resolve(htmlPath);
        fileUrl = `file://${absolutePath}`;
    }

    console.log(`Opening: ${fileUrl}`);

    await page.goto(fileUrl, {
        waitUntil: 'networkidle',
        timeout: 60000
    });

    // Wait for Markdeep to render the slides
    await page.waitForSelector('.slide', { timeout: 30000 });

    // Give extra time for MathJax, Mermaid and other async rendering to complete
    await page.waitForTimeout(4000);

    // Extract slide data
    const slideData = await page.evaluate(() => {
        const PT_PER_PX = 0.75;
        const PX_PER_IN = 96;

        // Helper functions
        const pxToInch = (px) => px / PX_PER_IN;
        const pxToPoints = (pxStr) => parseFloat(pxStr) * PT_PER_PX;

        const rgbToHex = (rgbStr) => {
            if (!rgbStr || rgbStr === 'rgba(0, 0, 0, 0)' || rgbStr === 'transparent') return 'FFFFFF';
            const match = rgbStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
            if (!match) return '000000';
            return match.slice(1).map(n => parseInt(n).toString(16).padStart(2, '0')).join('').toUpperCase();
        };

        // Get presentation dimensions from slide CSS
        const slideElement = document.querySelector('.slide');
        const slideStyle = window.getComputedStyle(slideElement);
        const slideWidth = parseFloat(slideStyle.width);
        const slideHeight = parseFloat(slideStyle.height);

        // Calculate aspect ratio
        const aspectRatio = slideWidth / slideHeight;

        // Get slide dimensions in inches (for PPTX)
        const slideWidthInches = 10; // Standard PPTX width
        const slideHeightInches = slideWidthInches / aspectRatio;

        const slides = document.querySelectorAll('.slide');
        const extractedSlides = [];
        let mathIdCounter = 0;
        let chartIdCounter = 0;

        slides.forEach((slide, slideIndex) => {
            const slideContent = slide.querySelector('.slide-content');
            if (!slideContent) return;

            const elements = [];
            const slideRect = slideContent.getBoundingClientRect();
            const parentRect = slide.getBoundingClientRect();

            // Scale factors to convert from rendered pixels to PPTX inches
            const scaleX = slideWidthInches / parentRect.width;
            const scaleY = slideHeightInches / parentRect.height;

            // Extract slide classes for special handling
            const slideClasses = Array.from(slide.classList);
            const isSmallText = slideClasses.includes('small-text');
            const isTinyText = slideClasses.includes('tiny-text');
            const isH1TitleSlide = slideClasses.includes('h1-title-slide');
            const isTwoColumn = slideClasses.includes('two-column');

            // Function to extract text with formatting
            function extractTextWithFormatting(element) {
                const runs = [];

                function processNode(node, baseStyle = {}) {
                    if (node.nodeType === Node.TEXT_NODE) {
                        const text = node.textContent;
                        if (text.trim()) {
                            runs.push({ text, options: { ...baseStyle } });
                        }
                    } else if (node.nodeType === Node.ELEMENT_NODE) {
                        const computed = window.getComputedStyle(node);
                        const newStyle = { ...baseStyle };

                        // Check for bold
                        if (computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600) {
                            newStyle.bold = true;
                        }

                        // Check for italic
                        if (computed.fontStyle === 'italic') {
                            newStyle.italic = true;
                        }

                        // Check for underline
                        if (computed.textDecoration && computed.textDecoration.includes('underline')) {
                            newStyle.underline = true;
                        }

                        // Check for color
                        if (computed.color && computed.color !== 'rgb(0, 0, 0)') {
                            newStyle.color = rgbToHex(computed.color);
                        }

                        // Handle line breaks
                        if (node.tagName === 'BR') {
                            runs.push({ text: '\n', options: {} });
                            return;
                        }

                        // Process children
                        node.childNodes.forEach(child => processNode(child, newStyle));
                    }
                }

                processNode(element);
                return runs;
            }

            // Like extractTextWithFormatting, but inserts a line break between top-level
            // <p> children so multi-paragraph quote bodies don't run together.
            function extractParagraphRuns(container) {
                const paragraphs = container.querySelectorAll(':scope > p');
                if (paragraphs.length === 0) {
                    return extractTextWithFormatting(container);
                }
                const runs = [];
                paragraphs.forEach((p, idx) => {
                    runs.push(...extractTextWithFormatting(p));
                    if (idx < paragraphs.length - 1) {
                        runs.push({ text: '\n', options: {} });
                    }
                });
                return runs;
            }

            // Function to convert element position to PPTX coordinates
            function getPosition(el) {
                const rect = el.getBoundingClientRect();
                return {
                    x: (rect.left - parentRect.left) * scaleX,
                    y: (rect.top - parentRect.top) * scaleY,
                    w: rect.width * scaleX,
                    h: rect.height * scaleY
                };
            }

            // Function to extract element styles
            function getElementStyle(el) {
                const computed = window.getComputedStyle(el);
                const fontSize = pxToPoints(computed.fontSize);

                // Apply text size modifiers
                let adjustedFontSize = fontSize;
                if (isSmallText) adjustedFontSize *= 0.85;
                if (isTinyText) adjustedFontSize *= 0.7;

                return {
                    fontSize: adjustedFontSize,
                    fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim() || 'Arial',
                    color: rgbToHex(computed.color),
                    bold: computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600,
                    italic: computed.fontStyle === 'italic',
                    underline: computed.textDecoration && computed.textDecoration.includes('underline'),
                    align: computed.textAlign === 'center' ? 'center' :
                        computed.textAlign === 'right' ? 'right' : 'left',
                    lineSpacing: computed.lineHeight !== 'normal' ? pxToPoints(computed.lineHeight) : null,
                    backgroundColor: computed.backgroundColor !== 'rgba(0, 0, 0, 0)' ? rgbToHex(computed.backgroundColor) : null
                };
            }

            // Process all content elements
            function processElement(el, depth = 0, inColumn = false) {
                const tagName = el.tagName;
                const position = getPosition(el);
                position.inColumn = inColumn;  // Mark if inside a column

                // Skip invisible or zero-size elements
                if (position.w === 0 || position.h === 0) return;

                // Skip anchor elements (navigation targets)
                if (tagName === 'A' && el.classList.contains('target')) return;

                // Handle Markdeep-specific title div (used on title slide)
                if (tagName === 'DIV' && el.classList.contains('title')) {
                    const style = getElementStyle(el);
                    style.bold = true;

                    elements.push({
                        type: 'heading',
                        level: 1,
                        text: extractTextWithFormatting(el),
                        position,
                        style
                    });
                    return;
                }

                // Handle Markdeep TOC title
                if (tagName === 'DIV' && el.classList.contains('toc-title')) {
                    const style = getElementStyle(el);
                    style.bold = true;

                    elements.push({
                        type: 'heading',
                        level: 1,
                        text: [{ text: el.textContent.trim(), options: { bold: true } }],
                        position,
                        style
                    });
                    return;
                }

                // Handle Markdeep subtitle (afterTitles div content)
                if (tagName === 'DIV' && el.classList.contains('afterTitles')) {
                    // Get the text after this div (subtitle text is a sibling text node)
                    const nextSibling = el.nextSibling;
                    if (nextSibling && nextSibling.nodeType === Node.TEXT_NODE) {
                        const subtitleText = nextSibling.textContent.trim();
                        if (subtitleText) {
                            const style = getElementStyle(el);
                            elements.push({
                                type: 'paragraph',
                                text: [{ text: subtitleText, options: {} }],
                                position: {
                                    x: position.x,
                                    y: position.y + 0.5,
                                    w: position.w,
                                    h: 0.5
                                },
                                style
                            });
                        }
                    }
                    return;
                }

                // Handle Markdeep TOC list
                if (tagName === 'UL' && el.classList.contains('toc-list')) {
                    const items = [];
                    const listItems = el.querySelectorAll(':scope > li');

                    listItems.forEach(li => {
                        const linkText = li.querySelector('a')?.textContent.trim() || li.textContent.trim();
                        items.push({
                            text: [{ text: linkText, options: {} }],
                            level: 0
                        });
                    });

                    elements.push({
                        type: 'list',
                        ordered: false,
                        items,
                        position,
                        style: getElementStyle(el)
                    });
                    return;
                }


                // Handle headings (H1-H6)
                if (/^H[1-6]$/.test(tagName)) {
                    const style = getElementStyle(el);
                    style.bold = true; // Headers are always bold
                    const level = parseInt(tagName[1]);

                    elements.push({
                        type: 'heading',
                        level,
                        text: extractTextWithFormatting(el),
                        position,
                        style
                    });
                    return;
                }

                // Handle paragraphs
                if (tagName === 'P') {
                    // If paragraph contains inline math, screenshot the whole paragraph
                    if (el.querySelector('.MathJax_SVG')) {
                        const mathId = `math-${mathIdCounter++}`;
                        el.setAttribute('data-math-id', mathId);
                        elements.push({ type: 'math_pending', mathId, position });
                        return;
                    }

                    // If paragraph contains an image, extract the image instead
                    const imgEl = el.querySelector('img');
                    if (imgEl) {
                        const imgRect = imgEl.getBoundingClientRect();
                        elements.push({
                            type: 'image',
                            src: imgEl.src,
                            alt: imgEl.alt || '',
                            position: {
                                x: (imgRect.left - parentRect.left) * scaleX,
                                y: (imgRect.top - parentRect.top) * scaleY,
                                w: imgRect.width * scaleX,
                                h: imgRect.height * scaleY,
                                inColumn
                            }
                        });
                        return;
                    }

                    const text = el.textContent.trim();
                    if (!text) return;

                    // Skip paragraphs that only contain a title element
                    if (el.querySelector('title')) return;

                    elements.push({
                        type: 'paragraph',
                        text: extractTextWithFormatting(el),
                        position,
                        style: getElementStyle(el)
                    });
                    return;
                }

                // Handle lists (UL, OL)
                if (tagName === 'UL' || tagName === 'OL') {
                    // If list contains inline math, screenshot the whole list
                    if (el.querySelector('.MathJax_SVG')) {
                        const mathId = `math-${mathIdCounter++}`;
                        el.setAttribute('data-math-id', mathId);
                        elements.push({ type: 'math_pending', mathId, position });
                        return;
                    }

                    const items = [];

                    // Recursive function to extract list items with nesting level
                    function extractListItems(listEl, level = 0) {
                        const listItems = listEl.querySelectorAll(':scope > li');

                        listItems.forEach(li => {
                            // Extract text from this LI (excluding nested lists)
                            const itemRuns = [];

                            li.childNodes.forEach(child => {
                                if (child.nodeType === Node.TEXT_NODE) {
                                    const text = child.textContent.trim();
                                    if (text) {
                                        itemRuns.push({ text, options: {} });
                                    }
                                } else if (child.nodeType === Node.ELEMENT_NODE) {
                                    // Skip nested UL/OL - will process separately
                                    if (child.tagName !== 'UL' && child.tagName !== 'OL') {
                                        const runs = extractTextWithFormatting(child);
                                        itemRuns.push(...runs);
                                    }
                                }
                            });

                            // Check for highlight classes
                            const highlightSpan = li.querySelector('[class*="highlight-"]');
                            if (highlightSpan) {
                                const highlightClass = Array.from(highlightSpan.classList).find(c => c.startsWith('highlight-'));
                                if (highlightClass) {
                                    const colorMap = {
                                        'highlight-red': 'C0392B',
                                        'highlight-orange': 'E67E22',
                                        'highlight-green': '27AE60',
                                        'highlight-blue': '2980B9',
                                        'highlight-purple': '8E44AD'
                                    };
                                    if (colorMap[highlightClass]) {
                                        itemRuns.forEach(run => run.options.highlightColor = colorMap[highlightClass]);
                                    }
                                }
                            }

                            // Only add if there's content
                            if (itemRuns.length > 0) {
                                items.push({
                                    text: itemRuns,
                                    level: level
                                });
                            }

                            // Process nested lists
                            const nestedLists = li.querySelectorAll(':scope > ul, :scope > ol');
                            nestedLists.forEach(nestedList => {
                                extractListItems(nestedList, level + 1);
                            });
                        });
                    }

                    extractListItems(el, 0);

                    elements.push({
                        type: 'list',
                        ordered: tagName === 'OL',
                        items,
                        position,
                        style: getElementStyle(el)
                    });
                    return;
                }

                // Handle images
                if (tagName === 'IMG') {
                    elements.push({
                        type: 'image',
                        src: el.src,
                        alt: el.alt || '',
                        position
                    });
                    return;
                }

                // Handle admonitions
                if (el.classList.contains('admonition')) {
                    const type = Array.from(el.classList).find(c => c !== 'admonition') || 'note';
                    const titleEl = el.querySelector('.admonitionTitle');
                    const title = titleEl ? titleEl.textContent.trim() : null;

                    // Get content (everything except title), preserving list formatting
                    const contentParts = [];

                    function extractAdmonitionContent(node) {
                        if (node === titleEl) return;

                        if (node.nodeType === Node.TEXT_NODE) {
                            const text = node.textContent.trim();
                            if (text) contentParts.push(text);
                        } else if (node.nodeType === Node.ELEMENT_NODE) {
                            if (node.tagName === 'UL' || node.tagName === 'OL') {
                                // Handle lists with bullet characters
                                const items = node.querySelectorAll(':scope > li');
                                items.forEach((li, idx) => {
                                    const prefix = node.tagName === 'OL' ? `${idx + 1}. ` : '• ';
                                    contentParts.push(prefix + li.textContent.trim());
                                });
                            } else if (node.tagName === 'P') {
                                const text = node.textContent.trim();
                                if (text) contentParts.push(text);
                            } else {
                                // Recurse for other elements
                                node.childNodes.forEach(child => extractAdmonitionContent(child));
                            }
                        }
                    }

                    el.childNodes.forEach(child => extractAdmonitionContent(child));

                    // Color mapping for admonition types
                    const colorMap = {
                        'note': { bg: 'E8F4FD', border: '2196F3', text: '1565C0' },
                        'tip': { bg: 'E8F5E9', border: '4CAF50', text: '2E7D32' },
                        'warning': { bg: 'FFF8E1', border: 'FFC107', text: 'F57F17' },
                        'error': { bg: 'FFEBEE', border: 'F44336', text: 'C62828' },
                        'question': { bg: 'FFF3E0', border: 'FF9800', text: 'E65100' }
                    };

                    const colors = colorMap[type] || colorMap['note'];

                    elements.push({
                        type: 'admonition',
                        admonitionType: type,
                        title,
                        content: contentParts.join('\n'),
                        position,
                        colors
                    });
                    return;
                }

                // Handle quote cards (:::quote 作者::: from markdeep-slides)
                if (el.classList.contains('quote-card')) {
                    const bodyEl = el.querySelector('.quote-body');
                    const authorEl = el.querySelector('.quote-author');
                    const bodyPos = bodyEl ? getPosition(bodyEl) : null;
                    const authorPos = authorEl ? getPosition(authorEl) : null;

                    elements.push({
                        type: 'quoteCard',
                        text: bodyEl ? extractParagraphRuns(bodyEl) : extractTextWithFormatting(el),
                        bodyOffsetY: bodyPos ? bodyPos.y - position.y : 0,
                        author: authorEl ? authorEl.textContent.trim() : null,
                        authorOffsetY: authorPos ? authorPos.y - position.y : null,
                        authorHeight: authorPos ? authorPos.h : 0,
                        authorStyle: authorEl ? getElementStyle(authorEl) : null,
                        position
                    });
                    return;
                }

                // Handle stat/KPI cards (:::stat::: from markdeep-slides)
                if (el.classList.contains('stat-card')) {
                    const numberEl = el.querySelector('.stat-number');
                    const labelEl = el.querySelector('.stat-label');
                    const numberPos = numberEl ? getPosition(numberEl) : null;
                    const labelPos = labelEl ? getPosition(labelEl) : null;

                    elements.push({
                        type: 'statCard',
                        number: numberEl ? numberEl.textContent.trim() : '',
                        numberOffsetY: numberPos ? numberPos.y - position.y : 0,
                        numberHeight: numberPos ? numberPos.h : position.h,
                        numberStyle: numberEl ? getElementStyle(numberEl) : null,
                        label: labelEl ? labelEl.textContent.trim() : '',
                        labelOffsetY: labelPos ? labelPos.y - position.y : 0,
                        labelHeight: labelPos ? labelPos.h : 0,
                        labelStyle: labelEl ? getElementStyle(labelEl) : null,
                        position
                    });
                    return;
                }

                // Handle code blocks
                if (tagName === 'PRE' || el.classList.contains('listing')) {
                    const codeEl = el.querySelector('code') || el;
                    elements.push({
                        type: 'code',
                        code: codeEl.textContent,
                        language: codeEl.className || '',
                        position,
                        style: getElementStyle(codeEl)
                    });
                    return;
                }

                // Handle tables
                if (tagName === 'TABLE') {
                    const rows = [];
                    el.querySelectorAll('tr').forEach(tr => {
                        const cells = [];
                        tr.querySelectorAll('th, td').forEach(cell => {
                            cells.push({
                                text: cell.textContent.trim(),
                                isHeader: cell.tagName === 'TH'
                            });
                        });
                        if (cells.length > 0) rows.push(cells);
                    });

                    elements.push({
                        type: 'table',
                        rows,
                        position,
                        style: getElementStyle(el)
                    });
                    return;
                }

                // Handle blockquotes (those not converted to admonitions)
                if (tagName === 'BLOCKQUOTE') {
                    elements.push({
                        type: 'blockquote',
                        text: extractTextWithFormatting(el),
                        position,
                        style: getElementStyle(el)
                    });
                    return;
                }

                // Handle columns container
                if (el.classList.contains('columns-container')) {
                    // Process columns separately
                    const leftCol = el.querySelector('.column-left');
                    const rightCol = el.querySelector('.column-right');

                    // Helper to add column background
                    function addColumnBackground(col) {
                        if (!col) return;
                        const colRect = col.getBoundingClientRect();
                        const colComputed = window.getComputedStyle(col);
                        const hasBg = colComputed.backgroundColor !== 'rgba(0, 0, 0, 0)';
                        const hasBorder = parseFloat(colComputed.borderWidth) > 0;

                        if (hasBg || hasBorder) {
                            elements.push({
                                type: 'shape',
                                isColumnBackground: true,
                                position: {
                                    x: (colRect.left - parentRect.left) * scaleX,
                                    y: (colRect.top - parentRect.top) * scaleY,
                                    w: colRect.width * scaleX,
                                    h: colRect.height * scaleY
                                },
                                fill: hasBg ? rgbToHex(colComputed.backgroundColor) : null,
                                border: hasBorder ? {
                                    color: rgbToHex(colComputed.borderColor),
                                    width: parseFloat(colComputed.borderWidth) * 0.75
                                } : null,
                                borderRadius: parseFloat(colComputed.borderRadius) || 0
                            });
                        }
                    }

                    // Add backgrounds first
                    addColumnBackground(leftCol);
                    addColumnBackground(rightCol);

                    if (leftCol) {
                        leftCol.childNodes.forEach(child => {
                            if (child.nodeType === Node.ELEMENT_NODE) {
                                processElement(child, depth + 1, true);  // Mark as inside column
                            }
                        });
                    }

                    if (rightCol) {
                        rightCol.childNodes.forEach(child => {
                            if (child.nodeType === Node.ELEMENT_NODE) {
                                processElement(child, depth + 1, true);  // Mark as inside column
                            }
                        });
                    }
                    return;
                }

                // Handle block-level MathJax (MathJax_SVG_Display)
                if (tagName === 'DIV' && el.classList.contains('MathJax_SVG_Display')) {
                    const mathId = `math-${mathIdCounter++}`;
                    el.setAttribute('data-math-id', mathId);
                    elements.push({ type: 'math_pending', mathId, position });
                    return;
                }

                // Handle Markdeep image container (div.image) — extract image + caption together
                if (tagName === 'DIV' && el.classList.contains('image')) {
                    const imgEl = el.querySelector('img');
                    if (imgEl) {
                        const captionEl = el.querySelector('span.imagecaption');
                        const caption = captionEl ? captionEl.textContent.trim() : null;
                        const imgRect = imgEl.getBoundingClientRect();
                        elements.push({
                            type: 'image',
                            src: imgEl.src,
                            alt: imgEl.alt || '',
                            caption,
                            position: {
                                x: (imgRect.left - parentRect.left) * scaleX,
                                y: (imgRect.top - parentRect.top) * scaleY,
                                w: imgRect.width * scaleX,
                                h: imgRect.height * scaleY,
                                inColumn
                            }
                        });
                    }
                    return;
                }

                // Handle inline SVG blocks — extract as vector SVG data URI (no screenshot needed)
                if (el.classList.contains('markdeep-svg')) {
                    const svgEl = el.querySelector('svg');
                    if (svgEl) {
                        try {
                            // ── Step 1: Resolve natural SVG dimensions from viewBox ──────────
                            // renderSVG() strips width/height attrs and sets style 100%×100%,
                            // so viewBox is the only reliable source of the intrinsic ratio.
                            let naturalW = 0, naturalH = 0;
                            const vbAttr = svgEl.getAttribute('viewBox');
                            if (vbAttr) {
                                const parts = vbAttr.trim().split(/[\s,]+/).map(parseFloat);
                                if (parts.length === 4 && parts[2] > 0 && parts[3] > 0) {
                                    naturalW = parts[2];  // viewBox width
                                    naturalH = parts[3];  // viewBox height
                                }
                            }

                            // ── Step 2: Compute PPTX size that fits inside the allocated
                            //           region while preserving the viewBox aspect ratio ──────
                            let svgW = position.w;
                            let svgH = position.h;
                            if (naturalW > 0 && naturalH > 0) {
                                const naturalRatio  = naturalW / naturalH;   // e.g. 1.0 for a square
                                const allocRatio    = position.w / position.h;
                                if (allocRatio > naturalRatio) {
                                    // allocated slot is wider → constrain by height
                                    svgW = position.h * naturalRatio;
                                    svgH = position.h;
                                } else {
                                    // allocated slot is taller → constrain by width
                                    svgW = position.w;
                                    svgH = position.w / naturalRatio;
                                }
                            }

                            // ── Step 3: Fix the SVG element before serialisation ─────────────
                            // Replace percentage-based style dims with the absolute viewBox
                            // values so PowerPoint/PptxGenJS know the intrinsic size and can
                            // scale without distortion.
                            svgEl.setAttribute('width',  naturalW  || svgW);
                            svgEl.setAttribute('height', naturalH  || svgH);
                            svgEl.style.width  = '';
                            svgEl.style.height = '';
                            // Make sure PowerPoint respects the aspect ratio when it scales the image
                            if (!svgEl.getAttribute('preserveAspectRatio')) {
                                svgEl.setAttribute('preserveAspectRatio', 'xMidYMid meet');
                            }

                            // ── Step 4: Serialise and push ────────────────────────────────────
                            const serializer = new XMLSerializer();
                            const svgStr = serializer.serializeToString(svgEl);
                            // btoa only handles latin1; encodeURIComponent+unescape handles UTF-8
                            const svgB64 = btoa(unescape(encodeURIComponent(svgStr)));
                            elements.push({
                                type: 'image',
                                src: 'data:image/svg+xml;base64,' + svgB64,
                                alt: 'svg',
                                position: {
                                    x: position.x,
                                    y: position.y,
                                    w: svgW,
                                    h: svgH,
                                    inColumn: position.inColumn
                                }
                            });
                        } catch (e) {
                            // Fallback: schedule a screenshot
                            const chartId = `chart-${chartIdCounter++}`;
                            el.setAttribute('data-chart-id', chartId);
                            elements.push({ type: 'chart_pending', chartId, position });
                        }
                    }
                    return;
                }

                // Handle Mermaid / ECharts / chart.js / D3 diagrams — screenshot as image
                if (el.classList.contains('markdeep-mermaid') ||
                    el.classList.contains('markdeep-echarts') ||
                    el.classList.contains('markdeep-chartjs') ||
                    el.classList.contains('markdeep-d3-force')) {
                    const chartId = `chart-${chartIdCounter++}`;
                    el.setAttribute('data-chart-id', chartId);
                    elements.push({ type: 'chart_pending', chartId, position });
                    return;
                }

                // Handle :::columns::: cells — the current markdeep-slides.js column markup
                // (columns-row > columns-cell), unlike the legacy column-left/column-right
                // layout below. Children must be flagged inColumn so list/paragraph rendering
                // picks the smaller in-card font and the card-padding Y-offset compensation.
                if (tagName === 'DIV' && el.classList.contains('columns-cell')) {
                    const cellComputed = window.getComputedStyle(el);
                    const cellHasBg = cellComputed.backgroundColor !== 'rgba(0, 0, 0, 0)';
                    const cellHasBorder = parseFloat(cellComputed.borderWidth) > 0;

                    if (cellHasBg || cellHasBorder) {
                        elements.push({
                            type: 'shape',
                            position,
                            fill: cellHasBg ? rgbToHex(cellComputed.backgroundColor) : null,
                            border: cellHasBorder ? {
                                color: rgbToHex(cellComputed.borderColor),
                                width: pxToPoints(cellComputed.borderWidth)
                            } : null,
                            borderRadius: parseFloat(cellComputed.borderRadius) || 0
                        });
                    }

                    el.childNodes.forEach(child => {
                        if (child.nodeType === Node.ELEMENT_NODE) {
                            processElement(child, depth + 1, true); // mark as inside column
                        }
                    });
                    return;
                }

                // Handle generic divs - process children
                if (tagName === 'DIV') {
                    // Check if this div has a background (shape)
                    const computed = window.getComputedStyle(el);
                    const hasBg = computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
                    const hasBorder = parseFloat(computed.borderWidth) > 0;

                    if (hasBg || hasBorder) {
                        elements.push({
                            type: 'shape',
                            position,
                            fill: hasBg ? rgbToHex(computed.backgroundColor) : null,
                            border: hasBorder ? {
                                color: rgbToHex(computed.borderColor),
                                width: pxToPoints(computed.borderWidth)
                            } : null,
                            borderRadius: parseFloat(computed.borderRadius) || 0
                        });
                    }

                    // Process children, preserving inColumn status
                    el.childNodes.forEach(child => {
                        if (child.nodeType === Node.ELEMENT_NODE) {
                            processElement(child, depth + 1, inColumn);
                        }
                    });
                    return;
                }

                // Fallback: recurse into any unrecognised container (CENTER, SECTION, FIGURE, A, etc.)
                // This ensures nested elements like images inside <center><a><img> are not lost.
                el.childNodes.forEach(child => {
                    if (child.nodeType === Node.ELEMENT_NODE) {
                        processElement(child, depth + 1, inColumn);
                    }
                });
            }

            // Process all direct children of slide content
            slideContent.childNodes.forEach(child => {
                if (child.nodeType === Node.ELEMENT_NODE) {
                    processElement(child);
                }
            });

            // Extract chapter label if present
            const chapterLabel = slide.querySelector('.chapter-label');

            // Extract slide number
            const slideNumber = slide.querySelector('.slide-number');

            // Extract navigation bar chapters
            const navItems = slide.querySelectorAll('.nav-section-item');
            let activeChapterIndex = -1;
            const navChapters = [];

            navItems.forEach((item, idx) => {
                const text = item.textContent.trim();
                // Skip TOC button
                if (!item.classList.contains('toc-button')) {
                    navChapters.push(text);
                    if (item.classList.contains('active')) {
                        activeChapterIndex = navChapters.length - 1;
                    }
                }
            });

            extractedSlides.push({
                index: slideIndex,
                id: slide.id,
                classes: slideClasses,
                elements,
                metadata: {
                    isH1TitleSlide,
                    isTwoColumn,
                    isSmallText,
                    isTinyText,
                    chapterLabel: chapterLabel ? chapterLabel.textContent.trim() : null,
                    slideNumber: slideNumber ? slideNumber.textContent.trim() : null,
                    navChapters: navChapters.length > 0 ? navChapters : null,
                    activeChapterIndex: activeChapterIndex >= 0 ? activeChapterIndex : null
                }
            });
        });

        // Get document title - try multiple sources
        const titleSlide = extractedSlides[0];
        let documentTitle = 'Untitled Presentation';

        // Try to get title from the first slide
        if (titleSlide && titleSlide.elements.length > 0) {
            // First, try heading elements
            const firstHeading = titleSlide.elements.find(e => e.type === 'heading');
            if (firstHeading && firstHeading.text && firstHeading.text.length > 0) {
                documentTitle = firstHeading.text.map(r => r.text).join('').trim();
            } else {
                // Try paragraph with bold text (Markdeep uses **text** which becomes <strong>)
                const firstParagraph = titleSlide.elements.find(e => e.type === 'paragraph');
                if (firstParagraph && firstParagraph.text && firstParagraph.text.length > 0) {
                    // Check if it contains bold text (likely the title)
                    const boldRun = firstParagraph.text.find(r => r.options?.bold);
                    if (boldRun) {
                        documentTitle = firstParagraph.text.map(r => r.text).join('').trim();
                    }
                }
            }
        }

        // Also try document.title as fallback
        if (documentTitle === 'Untitled Presentation') {
            const docTitle = document.title;
            if (docTitle && !docTitle.includes('localhost') && !docTitle.includes('file://')) {
                documentTitle = docTitle;
            }
        }

        // Try to extract from first STRONG element in the slide
        if (documentTitle === 'Untitled Presentation' && titleSlide) {
            const firstStrong = document.querySelector('.slide:first-of-type .slide-content strong');
            if (firstStrong) {
                documentTitle = firstStrong.textContent.trim();
            }
        }

        return {
            title: documentTitle,
            aspectRatio,
            dimensions: {
                width: slideWidthInches,
                height: slideHeightInches
            },
            slides: extractedSlides
        };
    });

    // Screenshot all math_pending elements and replace with image elements
    for (const slide of slideData.slides) {
        for (let i = 0; i < slide.elements.length; i++) {
            const el = slide.elements[i];
            if (el.type !== 'math_pending') continue;

            try {
                const locator = page.locator(`[data-math-id="${el.mathId}"]`);
                const pngBuf = await locator.screenshot({ type: 'png' });
                const dataUri = 'data:image/png;base64,' + pngBuf.toString('base64');
                slide.elements[i] = {
                    type: 'image',
                    src: dataUri,
                    alt: 'math',
                    position: el.position
                };
            } catch (err) {
                console.warn(`   ⚠ 数学公式截图失败 (${el.mathId}): ${err.message}`);
                slide.elements.splice(i, 1);
                i--;
            }
        }
    }

    // Screenshot all chart_pending elements and replace with image elements
    for (const slide of slideData.slides) {
        for (let i = 0; i < slide.elements.length; i++) {
            const el = slide.elements[i];
            if (el.type !== 'chart_pending') continue;

            try {
                const locator = page.locator(`[data-chart-id="${el.chartId}"]`);
                const pngBuf = await locator.screenshot({ type: 'png' });
                const dataUri = 'data:image/png;base64,' + pngBuf.toString('base64');
                slide.elements[i] = {
                    type: 'image',
                    src: dataUri,
                    alt: 'chart',
                    position: el.position
                };
                console.log(`   ✓ 图表截图成功 (${el.chartId})`);
            } catch (err) {
                console.warn(`   ⚠ 图表截图失败 (${el.chartId}): ${err.message}`);
                slide.elements.splice(i, 1);
                i--;
            }
        }
    }

    await browser.close();

    return slideData;
}

/**
 * Capture slide screenshots for reference
 * @param {string} htmlPath - Path to the HTML file
 * @param {string} outputDir - Directory to save screenshots
 */
export async function captureSlideScreenshots(htmlPath, outputDir) {
    const browser = await chromium.launch({ headless: true });
    const context = await browser.newContext({
        viewport: { width: 1920, height: 1080 }
    });

    const page = await context.newPage();

    let fileUrl = htmlPath;
    if (!htmlPath.startsWith('file://') && !htmlPath.startsWith('http')) {
        const absolutePath = path.resolve(htmlPath);
        fileUrl = `file://${absolutePath}`;
    }

    await page.goto(fileUrl, { waitUntil: 'networkidle' });
    await page.waitForSelector('.slide', { timeout: 30000 });
    await page.waitForTimeout(2000);

    const slideCount = await page.evaluate(() => document.querySelectorAll('.slide').length);

    const screenshots = [];

    for (let i = 0; i < slideCount; i++) {
        // Navigate to each slide
        await page.evaluate((index) => {
            if (typeof gotoSlide === 'function') {
                gotoSlide(index);
            }
        }, i);

        await page.waitForTimeout(500);

        const screenshotPath = path.join(outputDir, `slide-${i.toString().padStart(3, '0')}.png`);
        await page.screenshot({
            path: screenshotPath,
            fullPage: false,
            clip: await page.evaluate(() => {
                const slide = document.querySelector('.slide:not([style*="display: none"])');
                if (slide) {
                    const rect = slide.getBoundingClientRect();
                    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
                }
                return null;
            }) || undefined
        });

        screenshots.push(screenshotPath);
    }

    await browser.close();

    return screenshots;
}
