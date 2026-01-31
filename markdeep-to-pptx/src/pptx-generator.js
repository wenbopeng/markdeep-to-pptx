/**
 * PPTX Generator - Convert extracted slide data to PowerPoint presentations
 * 
 * This module takes the structured slide data extracted from Markdeep Slides
 * and generates a PowerPoint presentation using PptxGenJS.
 * 
 * Layout is based on the actual HTML element positions.
 */

import pptxgen from 'pptxgenjs';

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
    primary: '2980B9',          // Blue (from screenshots)
    titleText: '2980B9',        // Blue title
    bodyText: '333333',         // Dark gray
    lightText: '666666',        // Light gray
    white: 'FFFFFF',
    // Admonition colors
    noteBackground: 'E3F2FD',
    noteBorder: '2196F3',
    noteText: '1565C0',
    tipBackground: 'E8F5E9',
    tipBorder: '4CAF50',
    tipText: '2E7D32',
    warningBackground: 'FFF8E1',
    warningBorder: 'FFC107',
    warningText: 'F57F17',
    errorBackground: 'FFEBEE',
    errorBorder: 'F44336',
    errorText: 'C62828',
    questionBackground: 'FFF3E0',
    questionBorder: 'FF9800',
    questionText: 'E65100',
    // Bullet point color
    bulletColor: '2980B9'
};

// Standard slide dimensions
const SLIDE_WIDTH = 10;  // inches
const SLIDE_HEIGHT = 5.625; // 16:9

// Navigation bar height
const NAV_BAR_HEIGHT = 0.35;

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
        if (!isFirstSlide && !isH1TitleSlide && slideInfo.metadata?.navChapters) {
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
            renderContentSlide(slide, slideInfo, pptx);
        }

        // Add footer elements (chapter label and slide number)
        addFooter(slide, slideInfo, isFirstSlide, isH1TitleSlide);
    }

    await pptx.writeFile({ fileName: outputPath });
    console.log(`Presentation saved to: ${outputPath}`);
    return outputPath;
}

/**
 * Add footer with chapter label and slide number
 */
function addFooter(slide, slideInfo, isFirstSlide, isH1TitleSlide) {
    if (isFirstSlide || isH1TitleSlide) return;

    // Chapter label (bottom left)
    if (slideInfo.metadata?.chapterLabel) {
        slide.addText(slideInfo.metadata.chapterLabel, {
            x: 0.3,
            y: SLIDE_HEIGHT - 0.35,
            w: 3,
            h: 0.25,
            fontSize: FONT_SIZES.footer,
            color: COLORS.primary,
            fontFace: FONT_FACE
        });
    }

    // Slide number (bottom right) - wider to prevent line wrap
    if (slideInfo.metadata?.slideNumber) {
        slide.addText(slideInfo.metadata.slideNumber, {
            x: SLIDE_WIDTH - 1.2,
            y: SLIDE_HEIGHT - 0.35,
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

    // Calculate tab widths (distribute evenly)
    const tabWidth = SLIDE_WIDTH / chapters.length;

    chapters.forEach((chapter, idx) => {
        const x = idx * tabWidth;
        const isActive = idx === activeIndex;

        // Tab background (slightly lighter for active)
        if (isActive) {
            // Add bottom highlight for active tab
            slide.addShape(pptx.ShapeType.rect, {
                x: x,
                y: NAV_BAR_HEIGHT - 0.04,
                w: tabWidth,
                h: 0.04,
                fill: { color: COLORS.white },
                line: { type: 'none' }
            });
        }

        // Tab text
        slide.addText(chapter, {
            x: x,
            y: 0,
            w: tabWidth,
            h: NAV_BAR_HEIGHT,
            fontSize: 8,
            fontFace: FONT_FACE,
            color: COLORS.white,
            bold: isActive,
            align: 'center',
            valign: 'middle'
        });
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
        slide.addText(titleText, {
            x: 0.5,
            y: SLIDE_HEIGHT / 2 - 0.5,
            w: SLIDE_WIDTH - 1,
            h: 1,
            fontSize: FONT_SIZES.sectionTitle,
            fontFace: FONT_FACE,
            color: COLORS.primary,
            bold: true,
            align: 'center',
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
function renderContentSlide(slide, slideInfo, pptx) {
    // Find slide title (H2)
    const titleElement = slideInfo.elements.find(e => e.type === 'heading' && e.level === 2);

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
            fontSize: FONT_SIZES.slideTitle,
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
            case 'list':
                renderList(slide, element, pptx);
                break;
            case 'paragraph':
                renderParagraph(slide, element, pptx);
                break;
            case 'admonition':
                renderAdmonition(slide, element, pptx);
                break;
            case 'table':
                renderTable(slide, element, pptx);
                break;
            case 'code':
                renderCode(slide, element, pptx);
                break;
            case 'blockquote':
                renderBlockquote(slide, element, pptx);
                break;
        }
    }
}

/**
 * Render list using extracted position
 */
function renderList(slide, element, pptx) {
    const pos = element.position;
    if (!element.items || element.items.length === 0) return;

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
                    fontSize: FONT_SIZES.listItem,
                    bold: false
                }
            });
        } else {
            // Use bullet for unordered lists and nested items
            allTextRuns.push({
                text: `${indent}• `,
                options: {
                    color: COLORS.bulletColor,
                    fontSize: FONT_SIZES.listItem,
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
                    fontSize: FONT_SIZES.listItem
                }
            });
        });

        // Add line break after each item except last
        if (idx < element.items.length - 1) {
            allTextRuns.push({
                text: '\n',
                options: { fontSize: FONT_SIZES.listItem }
            });
        }
    });

    slide.addText(allTextRuns, {
        x: Math.max(pos.x, 0.5),
        y: Math.max(pos.y, 0.9),
        w: Math.min(pos.w, SLIDE_WIDTH - 1),
        h: Math.min(pos.h, SLIDE_HEIGHT - pos.y - 0.5),
        fontFace: FONT_FACE,
        valign: 'top',
        paraSpaceAfter: 8
    });
}

/**
 * Render paragraph
 */
function renderParagraph(slide, element, pptx) {
    const pos = element.position;
    const textRuns = formatTextRuns(element.text, FONT_SIZES.body);

    slide.addText(textRuns, {
        x: Math.max(pos.x, 0.5),
        y: pos.y,
        w: Math.min(pos.w, SLIDE_WIDTH - 1),
        h: Math.max(pos.h, 0.4),
        fontFace: FONT_FACE,
        valign: 'top'
    });
}

/**
 * Render admonition (callout box)
 */
function renderAdmonition(slide, element, pptx) {
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
        rectRadius: 0.03
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
    if (element.title) {
        slide.addText(element.title, {
            x: Math.max(pos.x, 0.5) + 0.2,
            y: contentY,
            w: Math.min(pos.w, SLIDE_WIDTH - 1) - 0.4,
            h: 0.3,
            fontSize: FONT_SIZES.smallText,
            fontFace: FONT_FACE,
            bold: true,
            color: colors.text
        });
        contentY += 0.35;
    }

    // Content
    if (element.content) {
        slide.addText(element.content, {
            x: Math.max(pos.x, 0.5) + 0.2,
            y: contentY,
            w: Math.min(pos.w, SLIDE_WIDTH - 1) - 0.4,
            h: height - (contentY - pos.y) - 0.1,
            fontSize: FONT_SIZES.smallText,
            fontFace: FONT_FACE,
            color: colors.text,
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

    const tableRows = element.rows.map((row) => {
        return row.map(cell => ({
            text: cell.text,
            options: {
                bold: cell.isHeader,
                fill: cell.isHeader ? COLORS.primary : 'F5F5F5',
                color: cell.isHeader ? 'FFFFFF' : COLORS.bodyText,
                fontSize: FONT_SIZES.smallText,
                align: 'center',
                valign: 'middle'
            }
        }));
    });

    const colCount = element.rows[0]?.length || 1;
    const tableWidth = Math.min(pos.w, SLIDE_WIDTH - 1);

    slide.addTable(tableRows, {
        x: Math.max(pos.x, 0.5),
        y: pos.y,
        w: tableWidth,
        colW: Array(colCount).fill(tableWidth / colCount),
        fontFace: FONT_FACE,
        border: { color: COLORS.lightText, pt: 0.5 }
    });
}

/**
 * Render code block
 */
function renderCode(slide, element, pptx) {
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
        fontSize: FONT_SIZES.code,
        fontFace: 'Courier New',
        color: '333333',
        valign: 'top'
    });
}

/**
 * Render blockquote
 */
function renderBlockquote(slide, element, pptx) {
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
        fontSize: FONT_SIZES.body,
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

    return runs.map(run => ({
        text: run.text || '',
        options: {
            bold: run.options?.bold,
            italic: run.options?.italic,
            underline: run.options?.underline ? { style: 'sng', color: COLORS.primary } : undefined,
            color: run.options?.bold ? COLORS.primary : COLORS.bodyText, // Bold text in blue
            fontSize: defaultSize
        }
    }));
}
