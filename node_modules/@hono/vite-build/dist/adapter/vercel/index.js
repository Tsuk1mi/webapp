import { existsSync, mkdirSync } from "node:fs";
import { cp, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import buildPlugin from "../../base.js";
const BUNDLE_NAME = "index.js";
const FUNCTION_NAME = "__hono";
const writeJSON = (path, data) => {
  const dir = resolve(path, "..");
  if (!existsSync(dir)) {
    mkdirSync(dir, { recursive: true });
  }
  return writeFile(path, JSON.stringify(data));
};
const getRuntimeVersion = () => {
  try {
    const systemNodeVersion = process.versions.node.split(".")[0];
    return `nodejs${Number(systemNodeVersion)}.x`;
  } catch {
    return "nodejs22.x";
  }
};
const vercelBuildPlugin = (pluginOptions) => {
  let config;
  return {
    ...buildPlugin({
      output: `functions/${FUNCTION_NAME}.func/${BUNDLE_NAME}`,
      outputDir: ".vercel/output",
      ...{
        entryContentAfterHooks: [
          // eslint-disable-next-line quotes
          () => "import { handle } from '@hono/node-server/vercel'"
        ],
        entryContentDefaultExportHook: (appName) => `export default handle(${appName})`
      },
      ...pluginOptions
    }),
    configResolved: (resolvedConfig) => {
      config = resolvedConfig;
    },
    writeBundle: async () => {
      const outputDir = resolve(config.root, config.build.outDir);
      const functionDir = resolve(outputDir, "functions", `${FUNCTION_NAME}.func`);
      const buildConfig = {
        ...pluginOptions?.vercel?.config,
        version: 3,
        routes: [
          ...pluginOptions?.vercel?.config?.routes ?? [],
          {
            handle: "filesystem"
          },
          {
            src: "/(.*)",
            dest: `/${FUNCTION_NAME}`
          }
        ]
      };
      const functionConfig = {
        ...pluginOptions?.vercel?.function,
        runtime: getRuntimeVersion(),
        launcherType: "Nodejs",
        handler: BUNDLE_NAME,
        shouldAddHelpers: true,
        shouldAddSourcemapSupport: Boolean(config.build.sourcemap),
        supportsResponseStreaming: true
      };
      await Promise.all([
        // Copy static files to the .vercel/output/static directory
        cp(resolve(config.root, config.publicDir), resolve(outputDir, "static"), {
          recursive: true
        }),
        // Write the all necessary config files
        writeJSON(resolve(outputDir, "config.json"), buildConfig),
        writeJSON(resolve(functionDir, ".vc-config.json"), functionConfig),
        writeJSON(resolve(functionDir, "package.json"), {
          type: "module"
        })
      ]);
    },
    name: "@hono/vite-build/vercel"
  };
};
var vercel_default = vercelBuildPlugin;
export {
  vercel_default as default
};
