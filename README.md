# Energy Monitor

A lightweight energy consumption dashboard. Drop in a `.docx` report and get an interactive view of savings — sites comparison, daily trend, and load mix. Word parsing happens entirely in the browser, no backend.

## Stack
- React 18 (UMD, no React build step)
- esbuild — JSX compilation and minification
- JSZip + DOMParser — read `.docx` in the browser
- Hand-rolled SVG charts with IntersectionObserver-driven animations

## Layout
```
.
├── index.html               # template, __BUILD__ → build hash
├── src/
│   ├── app.jsx              # all app logic
│   └── styles.css           # styles
├── public/
│   ├── favicon.svg
│   └── templates/           # sample .docx files for download
│       ├── template-campus.docx
│       └── template-mall.docx
├── build.mjs                # esbuild + public/ copy
├── package.json
├── vercel.json              # cache headers, output = dist/
└── dist/                    # build output (gitignored)
```

## Local
```bash
npm install
npm run build      # → dist/
npm run preview    # http://localhost:5173
```

## Deploy to Vercel

### Via CLI
```bash
npm i -g vercel
vercel              # first run — set up the project
vercel --prod       # production deploy
```
Vercel reads `vercel.json`, runs `npm run build`, and publishes `dist/`.

### Via GitHub
1. Push the repo to GitHub
2. On vercel.com → New Project → Import the repo
3. Framework Preset: **Other** (everything is configured in `vercel.json`)
4. Deploy

## What's optimized
- **No babel-standalone at runtime** — JSX is compiled at build time (–250 KB of download)
- **Build hashes in URLs** (`?v=abc123`) — immutable cache + automatic cache-bust on deploy
- **Cache-Control: immutable** for JS/CSS/SVG, `must-revalidate` for HTML
- **Pinned versions** of React/JSZip from CDN + `preconnect` for fast TLS
- **Minified** JS and CSS via esbuild
- **ASCII-only filenames** for templates (no encoded URLs)
- **Security headers**: X-Content-Type-Options, Referrer-Policy, X-Frame-Options, Permissions-Policy

## Input .docx format
The document should contain three sections with tables:
1. **Monthly summary** — Site | Period A | Period B | Saved %
2. **Daily breakdown** — two subsections, each with Day | Consumption
3. **Load mix** — two subsections, each with Type | Consumption | Share

The parser locates tables by shape (numeric columns, date format) rather than section names, so headings can vary. Site name is extracted from the H1 heading (e.g. `Energy report — Atlas Mall`) or from the filename. Both English and Russian source documents are supported.
