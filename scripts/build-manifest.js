/**
 * build-manifest.js
 * Generates production-ready manifest.xml and manifest.json for GitHub Pages deployment.
 *
 * Usage:
 *   node scripts/build-manifest.js --url https://<username>.github.io/<repo>/
 *   node scripts/build-manifest.js (reads from process.env.DEPLOY_URL or defaults to placeholder)
 */

const fs = require('fs');
const path = require('path');

function getArg(flag) {
  const index = process.argv.indexOf(flag);
  if (index !== -1 && index + 1 < process.argv.length) {
    return process.argv[index + 1];
  }
  return null;
}

const targetUrlArg = getArg('--url') || process.env.DEPLOY_URL || process.env.GH_PAGES_URL || 'https://<username>.github.io/<repo>/';
const targetUrl = targetUrlArg.endsWith('/') ? targetUrlArg : targetUrlArg + '/';
const targetOrigin = new URL(targetUrl).origin;

const projectRoot = path.join(__dirname, '..');
const manifestXmlPath = path.join(projectRoot, 'manifest.xml');
const distDir = path.join(projectRoot, 'dist');

if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, { recursive: true });
}

// 1. Process XML Manifest
if (fs.existsSync(manifestXmlPath)) {
  let xml = fs.readFileSync(manifestXmlPath, 'utf8');

  // Replace localhost references with production deployment base URL
  xml = xml.replace(/https:\/\/localhost:3000\//g, targetUrl);
  xml = xml.replace(/<AppDomain>https:\/\/localhost:3000<\/AppDomain>/g, `<AppDomain>${targetOrigin}</AppDomain>`);

  const distXmlPath = path.join(distDir, 'manifest.xml');
  fs.writeFileSync(distXmlPath, xml, 'utf8');
  console.log(`[build-manifest] Generated ${distXmlPath} -> Base URL: ${targetUrl}`);

  // Also write manifest.github.xml at root for convenient local sideloading / reference
  const rootGithubXmlPath = path.join(projectRoot, 'manifest.github.xml');
  fs.writeFileSync(rootGithubXmlPath, xml, 'utf8');
  console.log(`[build-manifest] Generated ${rootGithubXmlPath}`);
} else {
  console.warn(`[build-manifest] manifest.xml not found at ${manifestXmlPath}`);
}

// 2. Process JSON Manifest (Teams / Unified Manifest)
const distJsonPath = path.join(distDir, 'manifest.json');
if (fs.existsSync(distJsonPath)) {
  let jsonStr = fs.readFileSync(distJsonPath, 'utf8');
  jsonStr = jsonStr.replace(/https:\/\/localhost:3000\//g, targetUrl);
  jsonStr = jsonStr.replace(/"localhost:3000"/g, `"${new URL(targetUrl).host}"`);
  fs.writeFileSync(distJsonPath, jsonStr, 'utf8');
  console.log(`[build-manifest] Updated ${distJsonPath} with production URLs`);
}
