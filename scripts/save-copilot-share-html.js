#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

function safeSlug(url) {
  const tail = (url.split('/').pop() || 'copilot-share').split('?')[0];
  return tail.replace(/[^a-zA-Z0-9_-]/g, '') || 'copilot-share';
}

async function main() {
  const url = process.argv[2];
  const outputArg = process.argv[3];

  if (!url || !/^https:\/\/copilot\.microsoft\.com\/shares\//.test(url)) {
    console.error('Usage: node scripts/save-copilot-share-html.js <copilot-share-url> [output.html]');
    process.exit(1);
  }

  const outputFile = outputArg || `${safeSlug(url)}.html`;
  const outputPath = path.resolve(outputFile);

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 120000 });
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(1500);

    const html = await page.content();
    fs.writeFileSync(outputPath, html, 'utf8');
    console.log(`Saved rendered HTML to: ${outputPath}`);
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
