// Bundles src/web.ts + the docx library into docs/app.js. Run from a folder where `npm i docx esbuild` is installed.
import esbuild from "esbuild";
import fs from "node:fs";

// docx inlines a jszip copy whose setImmediate polyfill probes createElement("script"); these dead IE-era branches are cut.
const SCRIPT_PROBES = [
  [/"document" in \w+ && "onreadystatechange" in \w+\.document\.createElement\("script"\) \? function\(\) \{[\s\S]*?\} : (function\(\) \{)/g, "$1"],
  [/\w+ && "onreadystatechange" in \w+\.createElement\("script"\) \? \([\s\S]*?\}\) : (function\(\w+\) \{)/g, "$1"],
];
const noScriptProbe = {
  name: "no-script-probe",
  setup(build) {
    build.onLoad({ filter: /node_modules[\\/]docx[\\/]dist[\\/]index\.(mjs|cjs)$/ }, (args) => {
      let code = fs.readFileSync(args.path, "utf8");
      for (const [probe, keep] of SCRIPT_PROBES) {
        if (!probe.test(code)) throw new Error("setImmediate pattern not found in docx");
        code = code.replace(probe, keep);
      }
      return { contents: code, loader: "js" };
    });
  },
};
await esbuild.build({ entryPoints: ["src/web.ts"], bundle: true, plugins: [noScriptProbe], format: "iife", target: "es2020", platform: "browser", outfile: "docs/app.js", minify: false, logLevel: "info" });
