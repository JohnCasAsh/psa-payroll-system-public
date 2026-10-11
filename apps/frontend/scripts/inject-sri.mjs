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

// security.js and auth.css keep fixed names, so a browser holding an older copy
// would reuse it, and the old security.js then fails the integrity check (sign-in
// silently does nothing). A ?v=<content hash> makes every change a new address.
function versioned(file) {
  const assetPath = join(distDir, file);
  if (!existsSync(assetPath)) return file;
  return `${file}?v=${createHash("sha256").update(readFileSync(assetPath)).digest("hex").slice(0, 10)}`;
}

// Static pages referencing the shared, unhashed security.js and auth.css
for (const file of ["login.html", "register.html"]) {
  const path = join(distDir, file);
  if (!existsSync(path)) continue;
  let html = readFileSync(path, "utf-8");
  html = addIntegrity(html, "src", /<script[^>]*\ssrc="(security\.js)"[^>]*>/g);
  html = html.replace(/(\ssrc=")security\.js(")/g, `$1${versioned("security.js")}$2`);
  html = html.replace(/(\shref=")auth\.css(")/g, `$1${versioned("auth.css")}$2`);
  writeFileSync(path, html);
  console.log(`Injected SRI integrity and cache-busting versions into dist/${file}`);
}
