import { defineConfig } from "vite";
import devCerts from "office-addin-dev-certs";

export default defineConfig(async ({ command, isPreview }) => ({
  base: "./",
  server: {
    port: 3000,
    strictPort: true,
    https: command === "serve" && !isPreview ? await devCerts.getHttpsServerOptions() : undefined,
  },
  preview: {
    port: 3000,
    strictPort: true,
  },
  build: {
    target: "es2020",
  },
}));
