import { readFileSync, readdirSync } from "node:fs";
import { join, extname } from "node:path";

const ROOT = "src/pages";
const offenders = [];

function walk(dir) {
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, e.name);
    if (e.isDirectory()) walk(p);
    else if (extname(p) === ".html") checkFile(p);
  }
}

function checkFile(p) {
  const s = readFileSync(p, "utf8");
  const usesIncludes = /include/.test(s); // grov sjekk
  const bad = [/<header[\s>]/i, /<footer[\s>]/i, /<nav[\s>]/i, /<meta\s/i, /<link\s[^>]*stylesheet/i, /<script[^>]*app\.js/i];
  if (!usesIncludes && bad.some(rx => rx.test(s))) offenders.push(p);
}

walk(ROOT);

if (offenders.length) {
  console.error("Disse sidene ser ikke ut til å bruke partials:");
  offenders.forEach(f => console.error(" -", f));
  process.exit(1);
}
