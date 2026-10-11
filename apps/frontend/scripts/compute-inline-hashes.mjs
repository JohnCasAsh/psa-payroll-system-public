import { readFileSync } from "fs";
import { createHash } from "crypto";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const publicDir = join(__dirname, "..", "public");

function sha256(text) {
  return `sha256-${createHash("sha256").update(text, "utf-8").digest("base64")}`;
}

for (const file of ["login.html", "register.html"]) {
  const html = readFileSync(join(publicDir, file), "utf-8");

  const style = html.match(/<style>([\s\S]*?)<\/style>/);
  const inlineScript = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)]
    .find(m => !m[0].includes("src="));

  console.log(`\n${file}`);
  if (style) console.log(`  style-src  ${sha256(style[1])}`);
  if (inlineScript) console.log(`  script-src ${sha256(inlineScript[1])}`);
}
