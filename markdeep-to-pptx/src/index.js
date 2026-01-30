#!/usr/bin/env node

/**
 * Markdeep Slides to PPTX Converter
 * 
 * Main entry point for the converter.
 * 
 * Usage:
 *   node src/index.js <input.html> [output.pptx]
 *   
 * Examples:
 *   node src/index.js presentation.html
 *   node src/index.js presentation.html output/my-presentation.pptx
 */

import { extractSlides } from './slide-extractor.js';
import { generatePptx } from './pptx-generator.js';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function main() {
    const args = process.argv.slice(2);

    if (args.length === 0) {
        console.log(`
Markdeep Slides to PPTX Converter
=================================

Usage:
  node src/index.js <input.html> [output.pptx]

Arguments:
  input.html   - Path to the Markdeep Slides HTML file
  output.pptx  - Optional output path for the PPTX file (default: same name as input)

Examples:
  node src/index.js presentation.html
  node src/index.js ../markdeep-slides-project/Tutorial.html output/Tutorial.pptx
`);
        process.exit(0);
    }

    const inputPath = args[0];

    // Check if input file exists
    if (!fs.existsSync(inputPath)) {
        console.error(`Error: Input file not found: ${inputPath}`);
        process.exit(1);
    }

    // Determine output path
    let outputPath = args[1];
    if (!outputPath) {
        const inputBasename = path.basename(inputPath, path.extname(inputPath));
        const outputDir = path.join(path.dirname(__filename), '..', 'output');

        // Ensure output directory exists
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        outputPath = path.join(outputDir, `${inputBasename}.pptx`);
    }

    console.log(`
╔════════════════════════════════════════════════════════════════════╗
║           Markdeep Slides to PPTX Converter                        ║
╚════════════════════════════════════════════════════════════════════╝
`);

    console.log(`📄 Input:  ${inputPath}`);
    console.log(`📦 Output: ${outputPath}`);
    console.log('');

    try {
        // Step 1: Extract slides
        console.log('🔍 Step 1: Extracting slides from HTML...');
        const slideData = await extractSlides(inputPath);
        console.log(`   ✓ Extracted ${slideData.slides.length} slides`);
        console.log(`   ✓ Title: "${slideData.title}"`);
        console.log(`   ✓ Aspect ratio: ${slideData.aspectRatio.toFixed(2)}`);
        console.log('');

        // Step 2: Generate PPTX
        console.log('📊 Step 2: Generating PowerPoint presentation...');
        await generatePptx(slideData, outputPath);
        console.log(`   ✓ Presentation saved successfully`);
        console.log('');

        // Summary
        console.log('═══════════════════════════════════════════════════════════════════');
        console.log(`✅ Conversion complete!`);
        console.log(`   Open ${outputPath} to view your presentation.`);
        console.log('');

    } catch (error) {
        console.error('');
        console.error('❌ Error during conversion:');
        console.error(error.message);
        console.error('');
        console.error('Stack trace:');
        console.error(error.stack);
        process.exit(1);
    }
}

main();
