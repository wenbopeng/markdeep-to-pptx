#!/usr/bin/env node

/**
 * Debug script to extract navigation bar info
 */

import { chromium } from 'playwright';
import path from 'path';

async function debugNavbar() {
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

    // Extract navigation bar structure
    const navInfo = await page.evaluate(() => {
        // Find the navigation tabs
        const tabs = document.querySelectorAll('.tab');
        const chapters = [];

        tabs.forEach(tab => {
            chapters.push({
                text: tab.textContent.trim(),
                isActive: tab.classList.contains('active'),
                rect: tab.getBoundingClientRect()
            });
        });

        // Also check for navigation bar element
        const navBar = document.querySelector('.slideNavBar, .chapter-tabs, .nav-tabs, .tabs');

        return {
            tabs: chapters,
            navBarExists: !!navBar,
            tabCount: tabs.length
        };
    });

    console.log('\n=== NAVIGATION BAR INFO ===');
    console.log('Tab count:', navInfo.tabCount);
    console.log('Tabs:');
    navInfo.tabs.forEach((tab, idx) => {
        console.log(`  [${idx}] "${tab.text}" - Active: ${tab.isActive}`);
    });

    // Get the first slide's nav structure from a content slide
    await page.evaluate(() => {
        if (typeof gotoSlide === 'function') gotoSlide(3);
    });
    await page.waitForTimeout(500);

    const slideNavInfo = await page.evaluate(() => {
        const slide = document.querySelector('.slide:not([style*="display: none"])');
        if (!slide) return null;

        // Look for chapter tabs within or above the slide
        const tabs = document.querySelectorAll('.tab');
        const activeTab = document.querySelector('.tab.active');

        return {
            activeChaperTitle: activeTab?.textContent.trim(),
            allChapters: Array.from(tabs).map(t => t.textContent.trim())
        };
    });

    console.log('\n=== SLIDE 3 CHAPTER INFO ===');
    console.log('Active chapter:', slideNavInfo?.activeChaperTitle);
    console.log('All chapters:', slideNavInfo?.allChapters);

    await browser.close();
}

debugNavbar().catch(console.error);
