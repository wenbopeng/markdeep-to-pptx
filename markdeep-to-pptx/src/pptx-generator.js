/**
 * PPTX Generator - Convert extracted slide data to PowerPoint presentations
 * 
 * This module takes the structured slide data extracted from Markdeep Slides
 * and generates a PowerPoint presentation using PptxGenJS.
 * 
 * Layout is based on the actual HTML element positions.
 */

import pptxgen from 'pptxgenjs';
import fs from 'fs';
import https from 'https';
import http from 'http';
import path from 'path';
import { fileURLToPath } from 'url';

// Conversion factor: HTML pixels to PPTX inches (based on 1920px = 10 inches)
const PX_TO_INCH = 10 / 1920;

// Fixed font sizes for consistent output (in points)
const FONT_SIZES = {
    titleSlideTitle: 36,
    titleSlideSubtitle: 18,
    sectionTitle: 32,
    slideTitle: 24,
    body: 16,
    listItem: 16,
    smallText: 14,
    code: 12,
    footer: 9
};

// Font face - use Microsoft YaHei for Chinese text
const FONT_FACE = 'Microsoft YaHei';

// Color palette based on Markdeep default theme
const COLORS = {
    primary: '034295',          // Blue (from screenshots)
    titleText: '034295',        // Blue title
    bodyText: '333333',         // Dark gray
    lightText: '666666',        // Light gray
    white: 'FFFFFF',
    // Admonition colors
    noteBackground: 'd6eaf8',
    noteBorder: '7fb3d3',
    noteText: '1a5276',
    tipBackground: 'dcfad9',
    tipBorder: 'a2f29a',
    tipText: '1e8449',
    warningBackground: 'ffe9d5',
    warningBorder: 'ffc78f',
    warningText: 'a04000',
    errorBackground: 'fde8e8',
    errorBorder: 'f1948a',
    errorText: '922b21',
    questionBackground: 'f5eef8',
    questionBorder: '9b59b6',
    questionText: '6c3483',
    // Bullet point color
    bulletColor: '000000'
};

// Standard slide dimensions
const SLIDE_WIDTH = 10;  // inches
const SLIDE_HEIGHT = 5.625; // 16:9

// Navigation bar height
const NAV_BAR_HEIGHT = 0.35;

// Top padding offset inside column backgrounds (HTML padding-top → PPTX gap compensation)
const COLUMN_CONTENT_Y_OFFSET = 0.1;

/**
 * Create a PowerPoint presentation from extracted slide data
 */
export async function generatePptx(slideData, outputPath, options = {}) {
    const pptx = new pptxgen();

    // Set presentation metadata
    pptx.title = slideData.title || 'Markdeep Slides Presentation';
    pptx.author = options.author || 'Markdeep to PPTX Converter';

    // Set layout
    pptx.layout = 'LAYOUT_16x9';

    // Process each slide
    for (const slideInfo of slideData.slides) {
        const slide = pptx.addSlide();
        slide.background = { color: COLORS.white };

        const isFirstSlide = slideInfo.index === 0;
        const isH1TitleSlide = slideInfo.metadata?.isH1TitleSlide;
        const isTocSlide = slideInfo.elements.some(e =>
            e.type === 'heading' && e.text?.[0]?.text?.includes('目录'));

        // Add navigation bar for content slides (not first slide or section slides)
        if (!options.noNavbar && !isFirstSlide && !isH1TitleSlide && slideInfo.metadata?.navChapters) {
            renderNavBar(slide, slideInfo, pptx);
        }

        // Determine slide type and render accordingly
        if (isFirstSlide) {
            renderTitleSlide(slide, slideInfo, pptx);
        } else if (isH1TitleSlide) {
            renderSectionSlide(slide, slideInfo, pptx);
        } else if (isTocSlide) {
            renderTocSlide(slide, slideInfo, pptx);
        } else {
            await renderContentSlide(slide, slideInfo, pptx);
        }

        // Add footer elements (chapter label, slide number, and progress bar)
        addFooter(slide, slideInfo, isFirstSlide, isH1TitleSlide, pptx, slideData.slides.length, options);
    }

    await pptx.writeFile({ fileName: outputPath });
    console.log(`Presentation saved to: ${outputPath}`);
    return outputPath;
}

/**
 * Add footer with chapter label, slide number, and progress bar
 */
function addFooter(slide, slideInfo, isFirstSlide, isH1TitleSlide, pptx, totalSlides, options = {}) {
    // Progress bar at the very bottom (shown on ALL slides)
    if (!options.noProgressBar) {
        const progressBarHeight = 0.04;
        const currentSlide = slideInfo.index + 1;
        const progress = currentSlide / totalSlides;
        const progressWidth = SLIDE_WIDTH * progress;

        slide.addShape(pptx.ShapeType.rect, {
            x: 0,
            y: SLIDE_HEIGHT - progressBarHeight,
            w: progressWidth,
            h: progressBarHeight,
            fill: { color: COLORS.primary },
            line: { type: 'none' }
        });
    }

    // Skip chapter label and slide number for title/section slides
    if (isFirstSlide || isH1TitleSlide) return;

    // Chapter label (bottom left)
    if (!options.noChapter && slideInfo.metadata?.chapterLabel) {
        slide.addText(slideInfo.metadata.chapterLabel, {
            x: 0.3,
            y: SLIDE_HEIGHT - 0.4,
            w: 3,
            h: 0.25,
            fontSize: FONT_SIZES.footer,
            color: COLORS.primary,
            fontFace: FONT_FACE
        });
    }

    // Slide number (bottom right)
    if (!options.noPageNumber && slideInfo.metadata?.slideNumber) {
        slide.addText(slideInfo.metadata.slideNumber, {
            x: SLIDE_WIDTH - 1.2,
            y: SLIDE_HEIGHT - 0.4,
            w: 1,
            h: 0.25,
            fontSize: FONT_SIZES.footer,
            color: COLORS.lightText,
            fontFace: FONT_FACE,
            align: 'right'
        });
    }
}

/**
 * Render navigation bar at the top of the slide
 * Layout: [目录] on left (white bg), chapter tabs right-aligned with auto width
 * Active tab: inverted colors (white background, blue text)
 */
function renderNavBar(slide, slideInfo, pptx) {
    const chapters = slideInfo.metadata?.navChapters || [];
    const activeIndex = slideInfo.metadata?.activeChapterIndex;

    if (chapters.length === 0) return;

    // Background bar
    slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: SLIDE_WIDTH,
        h: NAV_BAR_HEIGHT,
        fill: { color: COLORS.primary },
        line: { type: 'none' }
    });

    // "目录" button on left (white background)
    const tocWidth = 0.5;
    slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: tocWidth,
        h: NAV_BAR_HEIGHT,
        fill: { color: COLORS.white },
        line: { type: 'none' }
    });
    slide.addText('目录', {
        x: 0,
        y: 0,
        w: tocWidth,
        h: NAV_BAR_HEIGHT,
        fontSize: 8,
        fontFace: FONT_FACE,
        color: COLORS.primary,
        align: 'center',
        valign: 'middle'
    });

    // Calculate tab widths based on text length (approximate)
    const charWidth = 0.12;  // Approximate width per Chinese character
    const tabPadding = 0.15;  // Padding on each side

    // Calculate each tab width
    const tabWidths = chapters.map(ch => ch.length * charWidth + tabPadding * 2);
    const totalTabWidth = tabWidths.reduce((a, b) => a + b, 0);

    // Start from right side
    let currentX = SLIDE_WIDTH - totalTabWidth;

    chapters.forEach((chapter, idx) => {
        const isActive = idx === activeIndex;
        const tabWidth = tabWidths[idx];

        // Tab background - white for active, transparent for others
        if (isActive) {
            slide.addShape(pptx.ShapeType.rect, {
                x: currentX,
                y: 0,
                w: tabWidth,
                h: NAV_BAR_HEIGHT,
                fill: { color: COLORS.white },
                line: { type: 'none' }
            });
        }

        // Tab text
        slide.addText(chapter, {
            x: currentX,
            y: 0,
            w: tabWidth,
            h: NAV_BAR_HEIGHT,
            fontSize: 8,
            fontFace: FONT_FACE,
            color: isActive ? COLORS.primary : COLORS.white,
            bold: false,
            align: 'center',
            valign: 'middle'
        });

        currentX += tabWidth;
    });
}

/**
 * Render title slide (first slide)
 */
function renderTitleSlide(slide, slideInfo, pptx) {
    const titleElement = slideInfo.elements.find(e => e.type === 'heading' && e.level === 1);
    const subtitleElement = slideInfo.elements.find(e => e.type === 'paragraph');

    // Center title
    const centerY = SLIDE_HEIGHT / 2 - 0.8;

    if (titleElement) {
        const titleText = extractPlainText(titleElement.text);
        slide.addText(titleText, {
            x: 0.5,
            y: centerY,
            w: SLIDE_WIDTH - 1,
            h: 1,
            fontSize: FONT_SIZES.titleSlideTitle,
            fontFace: FONT_FACE,
            color: COLORS.primary,
            bold: true,
            align: 'center',
            valign: 'middle'
        });
    }

    if (subtitleElement) {
        const subtitleText = extractPlainText(subtitleElement.text);
        slide.addText(subtitleText, {
            x: 0.5,
            y: centerY + 1.2,
            w: SLIDE_WIDTH - 1,
            h: 0.8,
            fontSize: FONT_SIZES.titleSlideSubtitle,
            fontFace: FONT_FACE,
            color: COLORS.lightText,
            align: 'center'
        });
    }
}

/**
 * Render section slide (H1 chapter transition)
 */
function renderSectionSlide(slide, slideInfo, pptx) {
    const titleElement = slideInfo.elements.find(e => e.type === 'heading' && e.level === 1);

    if (titleElement) {
        const titleText = extractPlainText(titleElement.text);
        const bannerH = 1.2;
        const bannerY = (SLIDE_HEIGHT - bannerH) / 2;

        // Full-width blue background banner
        slide.addShape(pptx.ShapeType.rect, {
            x: 0, y: bannerY,
            w: SLIDE_WIDTH, h: bannerH,
            fill: { color: COLORS.primary },
            line: { type: 'none' }
        });

        slide.addText(titleText, {
            x: 0.6,
            y: bannerY,
            w: SLIDE_WIDTH - 1.2,
            h: bannerH,
            fontSize: FONT_SIZES.sectionTitle,
            fontFace: FONT_FACE,
            color: COLORS.white,
            bold: true,
            align: 'left',
            valign: 'middle'
        });
    }
}

/**
 * Render TOC slide
 */
function renderTocSlide(slide, slideInfo, pptx) {
    // Check if this slide has navbar
    const hasNavBar = slideInfo.metadata?.navChapters?.length > 0;

    // Add navigation bar if present
    if (hasNavBar) {
        renderNavBar(slide, slideInfo, pptx);
    }

    const titleY = hasNavBar ? NAV_BAR_HEIGHT + 0.1 : 0.3;

    // Title - "目录" with underline
    const titleX = 0.5;
    const titleWidth = 2;

    slide.addText('目录', {
        x: titleX,
        y: titleY,
        w: titleWidth,
        h: 0.5,
        fontSize: FONT_SIZES.slideTitle,
        fontFace: FONT_FACE,
        color: COLORS.primary,
        bold: true
    });

    // TOC list - use text runs with explicit line breaks
    const listElement = slideInfo.elements.find(e => e.type === 'list');
    if (listElement && listElement.items) {
        const textRuns = [];

        listElement.items.forEach((item, idx) => {
            // Add number prefix
            textRuns.push({
                text: `${idx + 1}. `,
                options: {
                    fontSize: FONT_SIZES.body + 2,
                    color: COLORS.primary,
                    bold: true
                }
            });

            // Add item text
            textRuns.push({
                text: extractPlainText(item.text),
                options: {
                    fontSize: FONT_SIZES.body + 2,
                    color: COLORS.bodyText,
                    bold: false
                }
            });

            // Add line break after each item except last
            if (idx < listElement.items.length - 1) {
                textRuns.push({
                    text: '\n',
                    options: { fontSize: FONT_SIZES.body + 2 }
                });
            }
        });

        slide.addText(textRuns, {
            x: 1.5,
            y: titleY + 0.8,
            w: SLIDE_WIDTH - 3,
            h: SLIDE_HEIGHT - titleY - 1.5,
            fontFace: FONT_FACE,
            valign: 'top',
            lineSpacing: 32
        });
    }
}

/**
 * Render content slide (H2 title + content)
 * Uses actual positions from HTML extraction
 */
async function renderContentSlide(slide, slideInfo, pptx) {
    // Find slide title (H2)
    const titleElement = slideInfo.elements.find(e => e.type === 'heading' && e.level === 2);

    // Compute text scale for [small-text] / [tiny-text] slides
    const textScale = slideInfo.metadata?.isTinyText ? 0.7
                    : slideInfo.metadata?.isSmallText ? 0.85
                    : 1.0;

    // Check if this slide has navbar (to adjust title position)
    const hasNavBar = slideInfo.metadata?.navChapters?.length > 0;
    const titleY = hasNavBar ? NAV_BAR_HEIGHT + 0.1 : 0.3;

    // Add slide title using extracted position
    if (titleElement) {
        const pos = titleElement.position;
        const titleText = extractPlainText(titleElement.text);
        const titleWidth = Math.min(pos.w, SLIDE_WIDTH - 0.6);
        const titleX = Math.max(pos.x, 0.3);

        slide.addText(titleText, {
            x: titleX,
            y: titleY,
            w: titleWidth,
            h: 0.5,
            fontSize: Math.round(FONT_SIZES.slideTitle * textScale),
            fontFace: FONT_FACE,
            color: COLORS.titleText,
            bold: true
        });

        // Add underline below title
        const underlineY = titleY + 0.55;
        slide.addShape(pptx.ShapeType.rect, {
            x: titleX,
            y: underlineY,
            w: titleWidth,
            h: 0.025,
            fill: { color: COLORS.primary },
            line: { type: 'none' }
        });
    }

    // Process other content elements
    for (const element of slideInfo.elements) {
        if (element.type === 'heading' && element.level <= 2) continue; // Skip H1/H2

        switch (element.type) {
            case 'heading':
                // Render H3+ as section subheadings
                renderSubheading(slide, element, pptx, textScale);
                break;
            case 'list':
                renderList(slide, element, pptx, textScale);
                break;
            case 'paragraph':
                renderParagraph(slide, element, pptx, textScale);
                break;
            case 'admonition':
                renderAdmonition(slide, element, pptx, textScale);
                break;
            case 'table':
                renderTable(slide, element, pptx);
                break;
            case 'code':
                renderCode(slide, element, pptx, textScale);
                break;
            case 'blockquote':
                renderBlockquote(slide, element, pptx, textScale);
                break;
            case 'shape':
                renderShape(slide, element, pptx);
                break;
            case 'image':
                await renderImage(slide, element);
                break;
        }
    }
}

/**
 * Fetch a URL and return a Buffer
 */
function fetchBuffer(url) {
    return new Promise((resolve, reject) => {
        const client = url.startsWith('https') ? https : http;
        client.get(url, res => {
            const chunks = [];
            res.on('data', chunk => chunks.push(chunk));
            res.on('end', () => resolve(Buffer.concat(chunks)));
            res.on('error', reject);
        }).on('error', reject);
    });
}

/**
 * Render image element
 */
async function renderImage(slide, element) {
    const pos = element.position;
    const src = element.src || '';

    if (!src) return;

    try {
        let imageData;
        let extName;

        if (src.startsWith('data:image/')) {
            // Base64 data URI — pass directly
            const mimeMatch = src.match(/data:image\/([a-zA-Z+]+);base64,/);
            extName = mimeMatch ? mimeMatch[1].replace('jpeg', 'jpg').replace('svg+xml', 'svg') : 'png';
            imageData = src; // PptxGenJS accepts data URIs natively
        } else if (src.startsWith('http://') || src.startsWith('https://')) {
            // Remote URL — download to buffer
            const buf = await fetchBuffer(src);
            const urlPath = new URL(src).pathname;
            extName = path.extname(urlPath).replace('.', '') || 'png';
            imageData = `data:image/${extName};base64,` + buf.toString('base64');
        } else {
            // Local file path (file:// or relative)
            const filePath = src.startsWith('file://') ? fileURLToPath(src) : src;
            extName = path.extname(filePath).replace('.', '') || 'png';
            imageData = filePath;
        }

        // Normalise extension
        if (extName === 'jpg') extName = 'jpeg';

        const imgX = Math.max(pos.x, 0);
        const imgW = Math.min(pos.w, SLIDE_WIDTH);

        slide.addImage({
            data: imageData,
            x: imgX,
            y: pos.y,
            w: imgW,
            h: pos.h
        });

        if (element.caption) {
            slide.addText(element.caption, {
                x: imgX,
                y: pos.y + pos.h + 0.05,
                w: imgW,
                h: 0.25,
                fontSize: FONT_SIZES.footer,
                fontFace: FONT_FACE,
                color: COLORS.lightText,
                align: 'center',
                italic: true
            });
        }
    } catch (err) {
        console.warn(`   ⚠ 图片加载失败，已跳过: ${src.slice(0, 80)} (${err.message})`);
    }
}

/**
 * Render background shape (used for column backgrounds)
 */
function renderShape(slide, element, pptx) {
    const pos = element.position;

    // Compensate for HTML padding: top is absorbed by COLUMN_CONTENT_Y_OFFSET,
    // so shrink height by the same amount to tighten the bottom border too.
    const extraHeight = -COLUMN_CONTENT_Y_OFFSET;

    slide.addShape(pptx.ShapeType.roundRect, {
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h + extraHeight,
        fill: element.fill ? { color: element.fill } : { color: 'F5F5F5' },
        line: element.border ? {
            color: element.border.color,
            width: element.border.width
        } : { type: 'none' },
        rectRadius: 0.1
    });
}

/**
 * Render H3+ subheadings
 */
function renderSubheading(slide, element, pptx, textScale = 1.0) {
    const pos = element.position;
    const text = extractPlainText(element.text);

    // Font size decreases for deeper levels
    const fontSizeMap = {
        3: FONT_SIZES.body + 2,  // H3: slightly larger than body
        4: FONT_SIZES.body,      // H4: same as body
        5: FONT_SIZES.smallText, // H5: smaller
        6: FONT_SIZES.smallText  // H6: smaller
    };
    const fontSize = Math.round((fontSizeMap[element.level] || FONT_SIZES.body) * textScale);

    // Use text runs for proper bold styling
    slide.addText([{
        text: text,
        options: {
            fontSize: fontSize,
            fontFace: FONT_FACE,
            color: COLORS.primary,
            bold: true
        }
    }], {
        x: Math.max(pos.x, 0.5),
        y: pos.y,
        w: Math.min(pos.w, SLIDE_WIDTH - 1),
        h: 0.35,  // Fixed smaller height
        valign: 'top'
    });
}

/**
 * Render list using extracted position
 */
function renderList(slide, element, pptx, textScale = 1.0) {
    const pos = element.position;
    if (!element.items || element.items.length === 0) return;

    // Use smaller font for column content
    const fontSize = Math.round((pos.inColumn ? FONT_SIZES.smallText : FONT_SIZES.listItem) * textScale);

    // Build list items with explicit bullet characters and indentation
    const allTextRuns = [];
    const INDENT_SIZE = '    ';  // 4 spaces per level

    element.items.forEach((item, idx) => {
        const level = item.level || 0;
        const indent = INDENT_SIZE.repeat(level);

        // Add indentation + bullet character or number prefix
        if (element.ordered && level === 0) {
            // Only number top-level items in ordered lists
            const topLevelIndex = element.items.slice(0, idx + 1)
                .filter(i => (i.level || 0) === 0).length;
            allTextRuns.push({
                text: `${indent}${topLevelIndex}. `,
                options: {
                    color: COLORS.bulletColor,
                    fontSize: fontSize,
                    bold: false
                }
            });
        } else {
            // Use bullet for unordered lists and nested items
            allTextRuns.push({
                text: `${indent}• `,
                options: {
                    color: COLORS.bulletColor,
                    fontSize: fontSize,
                    bold: false
                }
            });
        }

        // Format each text run with bold/italic preserved
        item.text.forEach((run) => {
            allTextRuns.push({
                text: run.text || '',
                options: {
                    bold: run.options?.bold,
                    italic: run.options?.italic,
                    color: run.options?.bold ? COLORS.primary : COLORS.bodyText,
                    fontSize: fontSize
                }
            });
        });

        // Add line break after each item except last
        if (idx < element.items.length - 1) {
            allTextRuns.push({
                text: '\n',
                options: { fontSize: fontSize }
            });
        }
    });

    const listY = pos.inColumn
        ? Math.max(pos.y - COLUMN_CONTENT_Y_OFFSET, 0.9)
        : Math.max(pos.y, 0.9);
    slide.addText(allTextRuns, {
        x: Math.max(pos.x, 0.5),
        y: listY,
        w: Math.min(pos.w, SLIDE_WIDTH - 1),
        h: Math.min(pos.h, SLIDE_HEIGHT - pos.y - 0.5),
        fontFace: FONT_FACE,
        valign: 'top',
        paraSpaceAfter: 8,
        lineSpacingMultiple: pos.inColumn ? 1.2 : 1.5
    });
}

/**
 * Render paragraph
 */
function renderParagraph(slide, element, pptx, textScale = 1.0) {
    const pos = element.position;
    // Use smaller font for column content
    const fontSize = Math.round((pos.inColumn ? FONT_SIZES.smallText : FONT_SIZES.body) * textScale);
    const textRuns = formatTextRuns(element.text, fontSize);

    const paraY = pos.inColumn ? pos.y - COLUMN_CONTENT_Y_OFFSET : pos.y;
    slide.addText(textRuns, {
        x: Math.max(pos.x, 0.5),
        y: paraY,
        w: Math.min(pos.w, SLIDE_WIDTH - 1),
        h: Math.max(pos.h, 0.4),
        fontFace: FONT_FACE,
        valign: 'top',
        lineSpacingMultiple: pos.inColumn ? 1.2 : 1.5
    });
}

/**
 * Render admonition (callout box)
 */
function renderAdmonition(slide, element, pptx, textScale = 1.0) {
    const pos = element.position;

    // Get colors based on type
    const typeColors = {
        note: { bg: COLORS.noteBackground, border: COLORS.noteBorder, text: COLORS.noteText },
        tip: { bg: COLORS.tipBackground, border: COLORS.tipBorder, text: COLORS.tipText },
        warning: { bg: COLORS.warningBackground, border: COLORS.warningBorder, text: COLORS.warningText },
        error: { bg: COLORS.errorBackground, border: COLORS.errorBorder, text: COLORS.errorText },
        question: { bg: COLORS.questionBackground, border: COLORS.questionBorder, text: COLORS.questionText }
    };

    const colors = typeColors[element.admonitionType] || typeColors.note;
    const height = Math.max(pos.h, 0.8);

    // Background shape with rounded corners
    slide.addShape(pptx.ShapeType.roundRect, {
        x: Math.max(pos.x, 0.5),
        y: pos.y,
        w: Math.min(pos.w, SLIDE_WIDTH - 1),
        h: height,
        fill: { color: colors.bg },
        line: { type: 'none' },
        rectRadius: 0.1
    });

    // Left accent bar
    slide.addShape(pptx.ShapeType.rect, {
        x: Math.max(pos.x, 0.5),
        y: pos.y,
        w: 0.06,
        h: height,
        fill: { color: colors.border },
        line: { type: 'none' }
    });

    // Title (if present)
    let contentY = pos.y + 0.12;
    const titleX = Math.max(pos.x, 0.5) + 0.2;
    const titleW = Math.min(pos.w, SLIDE_WIDTH - 1) - 0.4;

    if (element.title) {
        // Use text runs format for proper bold styling
        slide.addText([{
            text: element.title,
            options: {
                fontSize: Math.round(FONT_SIZES.body * textScale),
                fontFace: FONT_FACE,
                bold: true,
                color: COLORS.bodyText
            }
        }], {
            x: titleX,
            y: contentY,
            w: titleW,
            h: 0.35,
            valign: 'top'
        });

        // Underline below title
        slide.addShape(pptx.ShapeType.rect, {
            x: titleX,
            y: contentY + 0.32,
            w: titleW * 0.3,  // Partial underline for better visual
            h: 0.02,
            fill: { color: colors.border },
            line: { type: 'none' }
        });

        contentY += 0.45;
    }

    // Content
    if (element.content) {
        slide.addText(element.content, {
            x: Math.max(pos.x, 0.5) + 0.2,
            y: contentY,
            w: Math.min(pos.w, SLIDE_WIDTH - 1) - 0.4,
            h: height - (contentY - pos.y) - 0.1,
            fontSize: Math.round(FONT_SIZES.smallText * textScale),
            fontFace: FONT_FACE,
            color: COLORS.bodyText,
            valign: 'top'
        });
    }
}

/**
 * Render table
 */
function renderTable(slide, element, pptx) {
    const pos = element.position;
    if (!element.rows || element.rows.length === 0) return;

    const noBorder  = { type: 'none' };
    const thickLine = { pt: 1.5, color: COLORS.primary };
    const thinLine  = { pt: 0.75, color: COLORS.primary };

    const lastRowIdx = element.rows.length - 1;

    const tableRows = element.rows.map((row, rowIdx) => {
        const isHeader  = row.some(c => c.isHeader);
        const isLastRow = rowIdx === lastRowIdx;

        // top border: thick on header row, none otherwise
        // bottom border: thin after header, thick on last data row, none otherwise
        const topBorder    = isHeader   ? thickLine : noBorder;
        const bottomBorder = isHeader   ? thinLine
                           : isLastRow  ? thickLine
                           : noBorder;

        return row.map(cell => ({
            text: cell.text,
            options: {
                bold: isHeader,
                fill: 'FFFFFF',
                color: isHeader ? COLORS.primary : COLORS.bodyText,
                fontSize: 12,
                align: 'center',
                valign: 'middle',
                border: [topBorder, noBorder, bottomBorder, noBorder]
            }
        }));
    });

    const colCount = element.rows[0]?.length || 1;
    const availableWidth = SLIDE_WIDTH - 1;

    // Estimate natural column widths based on longest cell content
    const TABLE_FONT_SIZE = 12; // pt
    const PT_TO_INCH = 1 / 72;
    const CELL_PADDING = 0.2; // inches, both sides combined

    function estimateCellWidth(text) {
        let w = 0;
        for (const ch of (text || '')) {
            // CJK characters are roughly full-width; Latin/digits are half-width
            w += /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(ch)
                ? TABLE_FONT_SIZE * PT_TO_INCH * 1.1
                : TABLE_FONT_SIZE * PT_TO_INCH * 0.65;
        }
        return w + CELL_PADDING;
    }

    const naturalColWidths = Array(colCount).fill(0).map((_, ci) => {
        let max = 0.5; // minimum column width
        element.rows.forEach(row => {
            const w = estimateCellWidth(row[ci]?.text || '');
            if (w > max) max = w;
        });
        return max;
    });

    const totalNatural = naturalColWidths.reduce((a, b) => a + b, 0);
    // Scale to fill available width; if content is wider, scale down to fit
    const scale = availableWidth / totalNatural;
    const colW = naturalColWidths.map(w => w * scale);
    const tableWidth = availableWidth;

    const tableX = (SLIDE_WIDTH - tableWidth) / 2;

    slide.addTable(tableRows, {
        x: tableX,
        y: pos.y,
        w: tableWidth,
        colW,
        fontFace: FONT_FACE,
        border: noBorder
    });
}

/**
 * Render code block
 */
function renderCode(slide, element, pptx, textScale = 1.0) {
    const pos = element.position;
    const height = Math.max(pos.h, 0.5);

    // Background
    slide.addShape(pptx.ShapeType.rect, {
        x: Math.max(pos.x, 0.5),
        y: pos.y,
        w: Math.min(pos.w, SLIDE_WIDTH - 1),
        h: height,
        fill: { color: 'F5F5F5' },
        line: { type: 'none' }
    });

    // Code text
    slide.addText(element.code, {
        x: Math.max(pos.x, 0.5) + 0.1,
        y: pos.y + 0.08,
        w: Math.min(pos.w, SLIDE_WIDTH - 1) - 0.2,
        h: height - 0.16,
        fontSize: Math.round(FONT_SIZES.code * textScale),
        fontFace: 'Courier New',
        color: '333333',
        valign: 'top'
    });
}

/**
 * Render blockquote
 */
function renderBlockquote(slide, element, pptx, textScale = 1.0) {
    const pos = element.position;
    const height = Math.max(pos.h, 0.4);

    // Left border
    slide.addShape(pptx.ShapeType.rect, {
        x: Math.max(pos.x, 0.5),
        y: pos.y,
        w: 0.04,
        h: height,
        fill: { color: COLORS.primary },
        line: { type: 'none' }
    });

    // Quote text
    const text = extractPlainText(element.text);
    slide.addText(text, {
        x: Math.max(pos.x, 0.5) + 0.15,
        y: pos.y,
        w: Math.min(pos.w, SLIDE_WIDTH - 1) - 0.15,
        h: height,
        fontSize: Math.round(FONT_SIZES.body * textScale),
        fontFace: 'Georgia',
        italic: true,
        color: COLORS.lightText,
        valign: 'top'
    });
}

// ============ Helper Functions ============

/**
 * Extract plain text from text runs
 */
function extractPlainText(runs) {
    if (!runs || !Array.isArray(runs)) return '';
    return runs.map(r => r.text || '').join('').trim();
}

/**
 * Format text runs for PptxGenJS with proper styling
 */
function formatTextRuns(runs, defaultSize) {
    if (!runs || !Array.isArray(runs)) return [{ text: '', options: {} }];

    // Strip leading newlines/whitespace from the very first non-empty run
    let firstNonEmpty = true;
    return runs.map(run => {
        let text = run.text || '';
        if (firstNonEmpty && text.trim()) {
            text = text.replace(/^[\n\r\s]+/, '');
            firstNonEmpty = false;
        }
        return {
        text,
        options: {
            bold: run.options?.bold,
            italic: run.options?.italic,
            underline: run.options?.underline ? { style: 'sng', color: COLORS.primary } : undefined,
            color: run.options?.bold ? COLORS.primary : COLORS.bodyText, // Bold text in blue
            fontSize: defaultSize
        }
        };
    });
}
