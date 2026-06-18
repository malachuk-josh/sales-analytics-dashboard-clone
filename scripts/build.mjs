import esbuild from "esbuild";
import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";

const rootDir = process.cwd();
const distDir = path.join(rootDir, "dist");
const assetsDir = path.join(distDir, "assets");
const publicDir = path.join(rootDir, "public");
const buildVersion = Date.now().toString();

await fs.rm(distDir, { recursive: true, force: true });
await fs.mkdir(assetsDir, { recursive: true });

const aliasPlugin = {
  name: "app-alias",
  setup(build) {
    build.onResolve({ filter: /^recharts$/ }, () => ({
      path: path.join(rootDir, "node_modules", "recharts", "lib", "index.js"),
    }));

    build.onResolve({ filter: /^@\// }, (args) => {
      const basePath = path.join(rootDir, "src", args.path.slice(2));
      const resolvedPath = [basePath, `${basePath}.jsx`, `${basePath}.js`].find((candidate) => fsSync.existsSync(candidate));
      return { path: resolvedPath || basePath };
    });
  },
};

await esbuild.build({
  entryPoints: ["src/main.jsx"],
  bundle: true,
  minify: true,
  sourcemap: false,
  splitting: false,
  format: "esm",
  target: ["es2020"],
  outfile: "dist/assets/index.js",
  plugins: [aliasPlugin],
  loader: {
    ".js": "jsx",
    ".jsx": "jsx",
    ".css": "css",
  },
  define: {
    "process.env.NODE_ENV": '"production"',
  },
  logLevel: "info",
});

const sourceHtml = await fs.readFile(path.join(rootDir, "index.html"), "utf8");
const html = sourceHtml
  .replace(
    /<script type="module" src="\/(?:src\/main\.jsx|assets\/index\.js(?:\?v=[^"]+)?)"><\/script>/,
    `<script type="module" src="/assets/index.js?v=${buildVersion}"></script>`
  )
  .replace(
    /<link rel="stylesheet" href="\/assets\/index\.css(?:\?v=[^"]+)?" \/>/,
    ""
  )
  .replace("</head>", `    <link rel="stylesheet" href="/assets/index.css?v=${buildVersion}" />\n  </head>`);

await fs.writeFile(path.join(distDir, "index.html"), html);

if (fsSync.existsSync(publicDir)) {
  await fs.cp(publicDir, distDir, { recursive: true });
}
