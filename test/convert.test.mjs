// Conversion tests: bundle src/exporter.ts for Node, convert fixtures, unzip the .docx and check the text.
import test, { before } from "node:test";
import assert from "node:assert/strict";
import esbuild from "esbuild";
import JSZip from "jszip";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

let exportToDocx;
before(async () => {
  const out = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "m2w-")), "exporter.mjs");
  await esbuild.build({ entryPoints: ["src/exporter.ts"], bundle: true, format: "esm", platform: "node", outfile: out, logLevel: "silent" });
  ({ exportToDocx } = await import(pathToFileURL(out).href));
});

const unescape = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

async function convert(md) {
  const bytes = await exportToDocx(md, { title: "t" });
  assert.equal(bytes[0], 0x50); // "PK"
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("word/document.xml").async("string");
  assert.ok(xml.trimEnd().endsWith("</w:document>"), "document.xml is complete");
  assert.ok(zip.file("[Content_Types].xml"), "content types present");
  const text = unescape([...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("\n"));
  return { xml, text, zip };
}

const fixtures = {
  headings: ["# One\n\n## Two\n\n###### Six\n\nSetext-free body", ["One", "Two", "Six", "Setext-free body"], (x) => /w:val="Heading1"/.test(x) && /w:val="Heading6"/.test(x)],
  "nested lists": ["- a\n  - b\n    - c\n- d\n\n1. x\n2. y\n   1. z\n", ["a", "b", "c", "d", "x", "y", "z"], (x) => /<w:ilvl w:val="2"/.test(x)],
  "tables": ["| H1 | H2 |\n|:--|--:|\n| a | b |\n| c | **d** |\n", ["H1", "H2", "a", "b", "c", "d"], (x) => /<w:tbl>/.test(x)],
  "table without outer pipes": ["A | B\n--- | ---\n1 | 2\n", ["A", "B", "1", "2"], (x) => /<w:tbl>/.test(x)],
  "code block": ["```js\nconst a = 1;\nif (a < 2 && b > 1) {}\n```\n", ["const a = 1;", "if (a < 2 && b > 1) {}"]],
  "tilde fence and empty code": ["~~~\nplain\n~~~\n\n```\n```\n\nafter", ["plain", "after"]],
  links: ["See [the docs](https://example.com/a?b=1&c=2) and https://example.org/x.\n", ["the docs", "https://example.org/x"], (x) => /<w:hyperlink/.test(x)],
  "image by URL": ["![logo](https://example.com/a.png)\n\ntext", ["logo", "text"]],
  "image local": ["![pic](images/a.png)\n\ntext", ["pic", "text"]],
  blockquote: ["> quoted line\n> second\n\nafter", ["quoted line", "second", "after"]],
  callout: ["> [!warning] Careful\n> body text\n", ["Careful", "body text"]],
  footnotes: ["Claim[^1] here.\n\n[^1]: The source.\n", ["Claim", "here."]],
  "task list": ["- [x] done\n- [ ] todo\n", ["done", "todo", "☑ ", "☐ "]],
  emoji: ["# Party 🎉\n\nHello 👋🏽 world 🚀 ❤️\n", ["Party 🎉", "Hello 👋🏽 world 🚀 ❤️"]],
  "CJK and Cyrillic": ["# 見出し\n\n日本語のテキスト、中文文本，한국어\n\nПривет, мир! Это тест.\n", ["見出し", "日本語のテキスト、中文文本，한국어", "Привет, мир! Это тест."]],
  "inline styles": ["**bold** *it* ~~gone~~ ==hi== `code` ***both***", ["bold", "it", "gone", "hi", "code", "both"], (x) => /<w:b\/>/.test(x) && /<w:strike\/>/.test(x)],
  "xml special chars": ["a < b & c > d \"q\" 'x'\n", ['a < b & c > d "q" \'x\'']],
  frontmatter: ["---\ntitle: X\n---\n\n# Body\n", ["Body"]],
  "CRLF input": ["# Title\r\n\r\n- a\r\n- b\r\n\r\ntext\r\n", ["Title", "a", "b", "text"]],
  "html and underscores": ["snake_case_name and 2*3*4 and <b>raw</b> text\n", ["snake_case_name", "raw"]],
  "hr and escapes": ["a\n\n---\n\nprice \\*not bold\\* 5 \\# x", ["a", "price *not bold* 5 # x"]],
  "multi-paragraph loose list": ["1. one\n\n2. two\n\n3. three\n", ["one", "two", "three"]],
  "bracket text": ["array[0] and [not a link] and (parens) [a](b)\n", ["array[0]", "[not a link]", "(parens)"]],
  "unterminated fence": ["```\nopen forever\n", ["open forever"]],
  "setext headings": ["Title\n=====\n\nSub\n---\n\nbody", ["Title", "Sub", "body"], (x) => /w:val="Heading1"/.test(x) && /w:val="Heading2"/.test(x), ["Title="]],
  "indented code": ["para\n\n    code line 1\n    code line 2\n\nafter", ["code line 1", "code line 2", "after"], (x) => /Consolas/.test(x)],
  "reference links": ["See [the docs][1] and [other] and [x][other].\n\n[1]: https://example.com\n[other]: https://example.org \"T\"\n", ["the docs", "other"], (x) => /<w:hyperlink/.test(x), ["[1]:", "[the docs][1]"]],
  autolinks: ["Visit <https://example.com/x> or <me@x.com>.", ["https://example.com/x", "me@x.com"], (x) => /<w:hyperlink/.test(x), ["&lt;", "<https"]],
  "html comments": ["a\n\n<!-- hidden -->\n\nb\n\n```html\n<!-- kept in code -->\n```\n", ["a", "b", "<!-- kept in code -->"], null, ["hidden"]],
  "html blocks": ["<div align=\"center\">\n<img src=\"a.png\" alt=\"Logo\">\n</div>\n\ntext with <kbd>Ctrl</kbd>", ["text with", "Ctrl"], null, ["<div", "<img", "</div>", "<kbd>"]],
  "badge links": ["[![Build](https://img.shields.io/b.svg)](https://ci.example.com)", ["Build"], (x) => /<w:hyperlink/.test(x), ["Build]"]],
  "backslash line break": ["line one\\\nline two\n", ["line one", "line two"], null, ["line one\\"]],
  "only whitespace-ish": ["\n\n   \n", []],
};

for (const [name, [md, expected, extra, absent = []]] of Object.entries(fixtures)) {
  test(`fixture: ${name}`, async () => {
    const { xml, text } = await convert(md);
    for (const e of expected) assert.ok(text.includes(e), `missing ${JSON.stringify(e)} in:\n${text}`);
    for (const a of absent) assert.ok(!text.includes(a), `unexpected ${JSON.stringify(a)} in:\n${text}`);
    if (extra) assert.ok(extra(xml), "structure check failed");
  });
}

test("5,000-line document converts quickly and keeps every line", async () => {
  const parts = [];
  for (let i = 0; i < 1000; i++) parts.push(`## Section ${i}`, "", `Paragraph ${i} with **bold** and a [link](https://example.com/${i}).`, "", `- item ${i}`, "- [ ] task " + i);
  const md = parts.join("\n");
  assert.ok(md.split("\n").length >= 5000);
  const t0 = Date.now();
  const { text } = await convert(md);
  const ms = Date.now() - t0;
  assert.ok(text.includes("Section 999") && text.includes("Paragraph 500") && text.includes("task 999"));
  assert.ok(ms < 15000, `took ${ms} ms`);
  console.log(`# 5000-line doc: ${ms} ms`);
});

test("pathological input does not hang (long emphasis runs, deep nesting)", async () => {
  const md = "*a ".repeat(3000) + "\n\n" + "_x".repeat(3000) + "\n\n" + "> ".repeat(60) + "deep\n\n" + "[".repeat(2000) + "\n";
  const t0 = Date.now();
  const { text } = await convert(md);
  assert.ok(text.includes("deep"));
  assert.ok(Date.now() - t0 < 15000, `took ${Date.now() - t0} ms`);
});

test("shipped bundle docs/app.js converts through the page glue", async () => {
  const handlers = {};
  const els = {};
  const el = (id) => (els[id] ??= { id, value: "", textContent: "", addEventListener: (ev, f) => ((handlers[id + ":" + ev] = f)), classList: { add() {}, remove() {} }, remove() {}, click() {} });
  let blob;
  const g = globalThis;
  const saved = { document: g.document, URL_c: URL.createObjectURL, URL_r: URL.revokeObjectURL, st: g.setTimeout };
  g.document = { getElementById: el, createElement: () => ({ click() { this.clicked = true; }, remove() {} }), body: { appendChild() {} } };
  URL.createObjectURL = (b) => ((blob = b), "blob:x");
  URL.revokeObjectURL = () => {};
  try {
    (0, eval)(fs.readFileSync("docs/app.js", "utf8"));
    el("md").value = "# Hi\n\nBody **text** with CJK 日本語";
    el("name").value = "my doc.docx";
    await handlers["convert:click"]();
    assert.match(els.status.textContent, /^Done: my doc\.docx/, els.status.textContent);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    assert.ok((await zip.file("word/document.xml").async("string")).includes("日本語"));
  } finally {
    g.document = saved.document;
    URL.createObjectURL = saved.URL_c;
    URL.revokeObjectURL = saved.URL_r;
  }
});
