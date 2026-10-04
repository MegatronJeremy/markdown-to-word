# Markdown to Word (.docx), in your browser
Static page (GitHub Pages serves the repo root; `docs/` holds an identical copy). Paste or drop Markdown, download a .docx. Runs entirely client-side; the page CSP has `connect-src 'none'`.

AI-assisted. MIT licence. The converter (`src/parser.ts`, `src/exporter.ts` and friends) is copied from the free DOCX Export Studio Obsidian plugin; it bundles the `docx` library (MIT).

Rebuild: `npm i docx esbuild && node build.mjs` writes `docs/app.js`; copy it and docs/index.html to the repo root.

There is an optional paid Pro of the plugin (not needed for this page): https://xparhyx.gumroad.com/l/bpfqja?utm_source=pages
