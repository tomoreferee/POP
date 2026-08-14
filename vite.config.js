import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Stamped in at build time so the version shown in the app is always the one
// actually running. Maintaining it by hand meant it drifted — it sat at an old
// date through dozens of changes and would have misled anyone debugging a
// "the fix isn't showing up" report.
const built = new Date();
const pad = (n) => String(n).padStart(2, "0");
const stamp = `${built.getFullYear()}-${pad(built.getMonth() + 1)}-${pad(built.getDate())} ${pad(built.getHours())}:${pad(built.getMinutes())}`;

export default defineConfig({
  plugins: [react()],
  define: {
    __APP_BUILD__: JSON.stringify(stamp),
  },
});
