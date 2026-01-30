#!/usr/bin/env node

/**
 * Debug script to see what's being extracted from slides
 */

import { extractSlides } from './slide-extractor.js';

async function debug() {
    const inputPath = process.argv[2] || '../markdeep-slides-project/20260120150333.html';

    console.log('Extracting slides from:', inputPath);

    const slideData = await extractSlides(inputPath);

    console.log('\n=== SLIDE DATA ===\n');
    console.log('Title:', slideData.title);
    console.log('Aspect Ratio:', slideData.aspectRatio);
    console.log('Total Slides:', slideData.slides.length);

    // Show slides 3-5 (content slides)
    for (let i = 3; i < Math.min(6, slideData.slides.length); i++) {
        const slide = slideData.slides[i];
        console.log(`\n--- Slide ${i} ---`);
        console.log('Classes:', slide.classes);
        console.log('Elements:');

        for (const el of slide.elements) {
            console.log(`  - Type: ${el.type}`);
            if (el.type === 'heading') {
                console.log(`    Level: ${el.level}`);
                console.log(`    Text:`, el.text?.map(r => r.text).join(''));
            } else if (el.type === 'paragraph') {
                const text = el.text?.map(r => r.text).join('');
                console.log(`    Text: ${text?.substring(0, 80)}...`);
            } else if (el.type === 'list') {
                console.log(`    Items: ${el.items?.length}`);
                el.items?.slice(0, 2).forEach((item, idx) => {
                    const itemText = item.text?.map(r => r.text).join('');
                    console.log(`      [${idx}]: ${itemText?.substring(0, 60)}...`);
                });
            } else if (el.type === 'admonition') {
                console.log(`    Type: ${el.admonitionType}`);
                console.log(`    Title: ${el.title}`);
            }
        }
    }
}

debug().catch(console.error);
