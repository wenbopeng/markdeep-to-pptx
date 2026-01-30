#!/usr/bin/env node

/**
 * Deep debug script to see raw HTML structure of slides
 */

import { chromium } from 'playwright';
import path from 'path';

async function debug() {
    const inputPath = process.argv[2] || '../markdeep-slides-project/20260120150333.html';

    const browser = await chromium.launch({ headless: true });
    const context = await browser.newContext({ viewport: { width: 1920, height: 1080 } });
    const page = await context.newPage();

    const absolutePath = path.resolve(inputPath);
    const fileUrl = `file://${absolutePath}`;

    console.log('Opening:', fileUrl);

    await page.goto(fileUrl, { waitUntil: 'networkidle', timeout: 60000 });
    await page.waitForSelector('.slide', { timeout: 30000 });
    await page.waitForTimeout(2000);

    // Get raw HTML structure of first 3 slides
    const slideHtml = await page.evaluate(() => {
        const slides = document.querySelectorAll('.slide');
        const result = [];

        for (let i = 0; i < Math.min(4, slides.length); i++) {
            const slide = slides[i];
            const slideContent = slide.querySelector('.slide-content');

            result.push({
                index: i,
                slideClasses: Array.from(slide.classList),
                contentHTML: slideContent ? slideContent.innerHTML.substring(0, 2000) : 'NO CONTENT',
                childTags: slideContent ? Array.from(slideContent.children).map(c => ({
                    tag: c.tagName,
                    class: c.className,
                    text: c.textContent?.substring(0, 100)
                })) : []
            });
        }

        return result;
    });

    console.log('\n=== RAW SLIDE HTML ===\n');
    for (const slide of slideHtml) {
        console.log(`\n--- Slide ${slide.index} ---`);
        console.log('Classes:', slide.slideClasses);
        console.log('\nChild elements:');
        for (const child of slide.childTags) {
            console.log(`  ${child.tag}.${child.class}: "${child.text?.substring(0, 80)}..."`);
        }
        console.log('\nRaw HTML (first 1000 chars):');
        console.log(slide.contentHTML.substring(0, 1000));
    }

    await browser.close();
}

debug().catch(console.error);
