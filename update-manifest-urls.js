#!/usr/bin/env node

/**
 * Script to update all localhost:3000 URLs in manifest.xml to production URL
 * Usage: node update-manifest-urls.js <production-url>
 * Example: node update-manifest-urls.js https://my-addin.vercel.app
 */

const fs = require('fs');
const path = require('path');

const manifestPath = path.join(__dirname, 'manifest.xml');
const productionUrl = process.argv[2];

if (!productionUrl) {
    console.error('Usage: node update-manifest-urls.js <production-url>');
    console.error('Example: node update-manifest-urls.js https://my-addin.vercel.app');
    process.exit(1);
}

// Ensure URL doesn't have trailing slash
const cleanUrl = productionUrl.replace(/\/$/, '');

// Read manifest
let manifest = fs.readFileSync(manifestPath, 'utf8');

// Replace all localhost:3000 URLs
const oldUrl = 'https://localhost:3000';
const newUrl = cleanUrl;

// Count replacements
const matches = manifest.match(new RegExp(oldUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'));
const count = matches ? matches.length : 0;

// Replace
manifest = manifest.replace(new RegExp(oldUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), newUrl);

// Write back
fs.writeFileSync(manifestPath, manifest, 'utf8');

console.log(`✅ Updated ${count} URLs in manifest.xml`);
console.log(`   ${oldUrl} → ${newUrl}`);
console.log(`\n📝 Next steps:`);
console.log(`   1. Update Azure redirect URI to: ${newUrl}/src/taskpane/taskpane.html`);
console.log(`   2. Update backend CORS to allow: ${newUrl}`);
console.log(`   3. Redeploy to Vercel`);
