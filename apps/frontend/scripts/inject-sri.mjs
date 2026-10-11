import { readFileSync, writeFileSync, existsSync } from "fs";
import { createHash } from "crypto";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const distDir = join(__dirname, "..", "dist");

function sriHash(filePath) {
  const buf = readFileSync(filePath);
  const digest = createHash("sha384").update(buf).digest("base64");
  return `sha384-${digest}`;
}

function addIntegrity(html, attr, urlPattern) {
  return html.replace(urlPattern, (tag, url) => {
    if (tag.includes("integrity=")) return tag;
    const assetPath = join(distDir, url.replace(/^\//, ""));
    if (!existsSync(assetPath)) return tag;
    const integrity = sriHash(assetPath);
    return tag.replace(new RegExp(`(${attr}="${url}")`), `$1 integrity="${integrity}"`);
  });
}

// Vite-built entry (dist/index.html): hashed /assets/*.js and *.css
{
  const path = join(distDir, "index.html");
  let html = readFileSync(path, "utf-8");
  html = addIntegrity(html, "src", /<script[^>]*\ssrc="(\/assets\/[^"]+\.js)"[^>]*>/g);
  html = addIntegrity(html, "href", /<link[^>]*\shref="(\/assets\/[^"]+\.css)"[^>]*>/g);
  writeFileSync(path, html);
  console.log("Injected SRI integrity attributes into dist/index.html");
}

// Static pages referencing the shared, unhashed security.js
for (const file of ["login.html", "register.html"]) {
  const path = join(distDir, file);
  if (!existsSync(path)) continue;
  let html = readFileSync(path, "utf-8");
  html = addIntegrity(html, "src", /<script[^>]*\ssrc="(security\.js)"[^>]*>/g);
  writeFileSync(path, html);
  console.log(`Injected SRI integrity attribute into dist/${file}`);
}
