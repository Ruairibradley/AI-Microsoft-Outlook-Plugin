import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

// Build directly into backend/web so the FastAPI server can host it.
// This is the key step to making everything "one localhost app".
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: path.resolve(__dirname, "../backend/web"),
    emptyOutDir: true
  }
});
