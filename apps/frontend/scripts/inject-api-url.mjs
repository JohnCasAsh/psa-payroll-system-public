import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const distDir = join(__dirname, "..", "dist");
const apiUrl = process.env.VITE_API_URL;

if (!apiUrl) {
  console.log("VITE_API_URL not set — skipping API URL injection (login/register.html will default to localhost).");
  process.exit(0);
}

// login/register set window._PSA_API_URL = '/api' (same-origin for local and
// LAN use). A hosted build (e.g. Firebase) points it at the real API instead.
// The older data-api-url marker is still handled for pages that use it.
const sameOrigin = /window\._PSA_API_URL\s*=\s*'\/api';/;
const legacyMarker = '<script src="security.js" data-api-url="">';
const safeUrl = JSON.stringify(apiUrl).slice(1, -1).replace(/'/g, "\\'");

for (const file of ["login.html", "register.html"]) {
  const path = join(distDir, file);
  const html = readFileSync(path, "utf-8");
  let injected = html;
  if (sameOrigin.test(html)) {
    injected = html.replace(sameOrigin, `window._PSA_API_URL = '${safeUrl}';`);
  } else if (html.includes(legacyMarker)) {
    injected = html.replace(legacyMarker, `<script src="security.js" data-api-url="${apiUrl}">`);
  } else {
    console.warn(`No API URL setting found in ${file}, skipping.`);
    continue;
  }
  writeFileSync(path, injected);
  console.log(`Set API URL to ${apiUrl} in dist/${file}`);
}
