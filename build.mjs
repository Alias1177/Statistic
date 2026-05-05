// Build script — produces dist/ ready for Vercel.
// Inputs:
//   - src/app.jsx       → dist/app.js   (esbuild: JSX → JS, minified)
//   - src/styles.css    → dist/styles.css (minified)
//   - index.html        → dist/index.html (with __BUILD__ replaced by hash)
//   - public/**         → dist/**        (copied as-is)
import { build } from "esbuild";
import { mkdir, copyFile, readFile, writeFile, readdir, stat, rm } from "node:fs/promises";
import { existsSync } from "node:fs";
import { join, dirname, relative } from "node:path";
import { createHash } from "node:crypto";

const ROOT = process.cwd();
const SRC  = join(ROOT, "src");
const PUB  = join(ROOT, "public");
const OUT  = join(ROOT, "dist");

async function copyTree(srcDir, outDir) {
  if (!existsSync(srcDir)) return;
  const entries = await readdir(srcDir, { withFileTypes: true });
  await mkdir(outDir, { recursive: true });
  for (const e of entries) {
    const s = join(srcDir, e.name);
    const o = join(outDir, e.name);
    if (e.isDirectory()) await copyTree(s, o);
    else await copyFile(s, o);
  }
}

async function main() {
  console.log("→ Cleaning dist/");
  try {
    await rm(OUT, { recursive: true, force: true });
  } catch (e) {
    // EPERM/EBUSY on some filesystems — fall back to overwriting in place
    if (!["EPERM", "EBUSY", "EACCES"].includes(e.code)) throw e;
    console.log("   (couldn't remove dist/ — will overwrite contents)");
  }
  await mkdir(OUT, { recursive: true });

  console.log("→ Bundling src/app.jsx (esbuild, minified)");
  await build({
    entryPoints: [join(SRC, "app.jsx")],
    bundle: false,        // We don't bundle — React, ReactDOM, JSZip stay as globals from CDN
    minify: true,
    target: ["es2019"],
    loader: { ".jsx": "jsx" },
    jsx: "transform",     // classic transform → uses global React.createElement
    outfile: join(OUT, "app.js"),
    logLevel: "info",
    legalComments: "none",
  });

  console.log("→ Minifying styles.css");
  await build({
    entryPoints: [join(SRC, "styles.css")],
    minify: true,
    outfile: join(OUT, "styles.css"),
    loader: { ".css": "css" },
    logLevel: "info",
    legalComments: "none",
  });

  console.log("→ Copying public/ → dist/");
  await copyTree(PUB, OUT);

  console.log("→ Processing index.html (cache-bust)");
  const appBuf = await readFile(join(OUT, "app.js"));
  const cssBuf = await readFile(join(OUT, "styles.css"));
  const hash = createHash("sha256")
    .update(appBuf).update(cssBuf)
    .digest("hex").slice(0, 10);
  let html = await readFile(join(ROOT, "index.html"), "utf8");
  html = html.replaceAll("__BUILD__", hash);
  await writeFile(join(OUT, "index.html"), html, "utf8");

  console.log("→ Building standalone.html (inlined CSS+JS, opens via file://)");
  // Standalone: inline CSS and JS into one self-contained HTML.
  // Templates won't be loadable via file:// (browser blocks fetch from file://),
  // so we hide the sample-download buttons in standalone mode by toggling a class.
  const cssText = cssBuf.toString("utf8");
  const jsText  = appBuf.toString("utf8");
  const standalone = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Energy Monitor · standalone</title>
  <style>${cssText}
.standalone .upload-btn.secondary { display: none; }
</style>
</head>
<body class="standalone">
  <div id="root"></div>
  <script src="https://unpkg.com/react@18.3.1/umd/react.production.min.js" crossorigin></script>
  <script src="https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js" crossorigin></script>
  <script src="https://unpkg.com/jszip@3.10.1/dist/jszip.min.js" crossorigin></script>
  <script>${jsText}</script>
</body>
</html>
`;
  await writeFile(join(OUT, "standalone.html"), standalone, "utf8");
  // Also drop a copy at the project root for easy double-click access
  try { await writeFile(join(ROOT, "open-me.html"), standalone, "utf8"); } catch {}

  console.log("→ Sizes:");
  for (const name of ["app.js", "styles.css", "index.html", "standalone.html"]) {
    const p = join(OUT, name);
    const s = await stat(p);
    console.log(`   ${name.padEnd(16)} ${(s.size / 1024).toFixed(1).padStart(7)} KB`);
  }
  console.log(`✓ Build hash: ${hash}`);
  console.log("✓ Output: dist/  (and open-me.html at project root)");
}

main().catch(e => { console.error(e); process.exit(1); });
