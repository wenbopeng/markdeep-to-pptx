#!/usr/bin/env node

/**
 * Visual debug script - capture screenshots and extract detailed layout info
 */

import { chromium } from 'playwright';
import path from 'path';
import fs from 'fs';

async function captureLayout() {
    const inputPath = process.argv[2] || '../markdeep-slides-project/20260120150333.html';
    const outputDir = './output/screenshots';

    // Create output directory
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    const browser = await chromium.launch({ headless: true });
    const context = await browser.newContext({ viewport: { width: 1920, height: 1080 } });
    const page = await context.newPage();

    const absolutePath = path.resolve(inputPath);
    const fileUrl = `file://${absolutePath}`;

    console.log('Opening:', fileUrl);

    await page.goto(fileUrl, { waitUntil: 'networkidle', timeout: 60000 });
    await page.waitForSelector('.slide', { timeout: 30000 });
    await page.waitForTimeout(3000);

    // Get slide count
    const slideCount = await page.evaluate(() => document.querySelectorAll('.slide').length);
    console.log(`Total slides: ${slideCount}\n`);

    // Analyze specific slides (content slides with H2)
    const slidesToAnalyze = [0, 3, 4, 5, 6];

    for (const slideIdx of slidesToAnalyze) {
        console.log(`\n========== SLIDE ${slideIdx} ==========`);

        // Navigate to slide
        await page.evaluate((idx) => {
            if (typeof gotoSlide === 'function') {
                gotoSlide(idx);
            }
        }, slideIdx);
        await page.waitForTimeout(500);

        // Capture screenshot
        const screenshotPath = `${outputDir}/slide-${slideIdx}.png`;
        await page.screenshot({ path: screenshotPath, fullPage: false });
        console.log(`Screenshot saved: ${screenshotPath}`);

        // Extract detailed layout info
        const layoutInfo = await page.evaluate((idx) => {
            const slides = document.querySelectorAll('.slide');
            const slide = slides[idx];
            if (!slide) return null;

            const slideRect = slide.getBoundingClientRect();
            const slideContent = slide.querySelector('.slide-content');
            if (!slideContent) return null;

            const contentRect = slideContent.getBoundingClientRect();

            const info = {
                slideWidth: slideRect.width,
                slideHeight: slideRect.height,
                contentPadding: {
                    left: contentRect.left - slideRect.left,
                    top: contentRect.top - slideRect.top,
                    right: slideRect.right - contentRect.right,
                    bottom: slideRect.bottom - contentRect.bottom
                },
                elements: []
            };

            // Get all visible elements and their positions
            function getElementInfo(el, depth = 0) {
                const rect = el.getBoundingClientRect();
                const computed = window.getComputedStyle(el);

                // Relative to slide content
                const relativeTop = rect.top - contentRect.top;
                const relativeLeft = rect.left - contentRect.left;

                return {
                    tag: el.tagName,
                    className: el.className,
                    text: el.textContent?.substring(0, 50),
                    position: {
                        top: relativeTop,
                        left: relativeLeft,
                        width: rect.width,
                        height: rect.height
                    },
                    style: {
                        fontSize: computed.fontSize,
                        fontWeight: computed.fontWeight,
                        color: computed.color,
                        marginTop: computed.marginTop,
                        marginBottom: computed.marginBottom,
                        paddingTop: computed.paddingTop,
                        paddingLeft: computed.paddingLeft
                    }
                };
            }

            // Analyze specific element types
            const h1 = slideContent.querySelector('h1');
            const h2 = slideContent.querySelector('h2');
            const lists = slideContent.querySelectorAll('ul, ol');
            const admonitions = slideContent.querySelectorAll('.admonition');

            if (h1) {
                info.elements.push({ type: 'H1', ...getElementInfo(h1) });
            }
            if (h2) {
                info.elements.push({ type: 'H2', ...getElementInfo(h2) });
            }

            lists.forEach((list, i) => {
                info.elements.push({ type: 'LIST', index: i, ...getElementInfo(list) });

                // Get first few list items
                const items = list.querySelectorAll(':scope > li');
                items.forEach((li, j) => {
                    if (j < 2) {
                        const liInfo = getElementInfo(li);
                        // Check for bullet/marker style
                        const marker = window.getComputedStyle(li, '::marker');
                        info.elements.push({
                            type: 'LI',
                            listIndex: i,
                            itemIndex: j,
                            listStyleType: window.getComputedStyle(list).listStyleType,
                            ...liInfo
                        });
                    }
                });
            });

            admonitions.forEach((adm, i) => {
                const admInfo = getElementInfo(adm);
                const titleEl = adm.querySelector('.admonitionTitle');
                info.elements.push({
                    type: 'ADMONITION',
                    index: i,
                    admonitionType: Array.from(adm.classList).find(c => c !== 'admonition'),
                    hasTitle: !!titleEl,
                    titleText: titleEl?.textContent,
                    ...admInfo
                });
            });

            return info;
        }, slideIdx);

        if (layoutInfo) {
            console.log(`\nSlide dimensions: ${layoutInfo.slideWidth}x${layoutInfo.slideHeight}`);
            console.log(`Content padding: L=${layoutInfo.contentPadding.left.toFixed(0)} T=${layoutInfo.contentPadding.top.toFixed(0)} R=${layoutInfo.contentPadding.right.toFixed(0)} B=${layoutInfo.contentPadding.bottom.toFixed(0)}`);

            console.log('\nElements:');
            for (const el of layoutInfo.elements) {
                console.log(`  ${el.type}${el.index !== undefined ? `[${el.index}]` : ''}:`);
                console.log(`    Position: top=${el.position.top.toFixed(0)}px, left=${el.position.left.toFixed(0)}px`);
                console.log(`    Size: ${el.position.width.toFixed(0)}x${el.position.height.toFixed(0)}`);
                console.log(`    Font: ${el.style.fontSize}, weight=${el.style.fontWeight}`);
                if (el.listStyleType) {
                    console.log(`    List style: ${el.listStyleType}`);
                }
                if (el.admonitionType) {
                    console.log(`    Admonition type: ${el.admonitionType}`);
                }
                console.log(`    Text: "${el.text?.substring(0, 40)}..."`);
            }
        }
    }

    await browser.close();
    console.log('\n\nDone! Check screenshots in:', outputDir);
}

captureLayout().catch(console.error);
