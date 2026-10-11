import { defineConfig, loadEnv } from "vite";

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");

  // Where the API actually runs. The browser never talks to this directly in dev —
  // it calls /api on the page's own origin and Vite proxies it here. That keeps every
  // request same-origin, so the app works from localhost, the LAN IP, or a phone
  // without listing each address in CSP or CORS.
  const apiTarget = env.VITE_API_TARGET ?? "http://127.0.0.1:3000";

  const devCsp = [
    "default-src 'self'",
    "script-src 'self' 'unsafe-eval' 'unsafe-inline'",
    "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
    "font-src 'self' https://fonts.gstatic.com",
    // 201 files are served through the /api proxy, so 'self' covers them. blob: stays
    // for anything generated in the page (exports, previews built client-side).
    "img-src 'self' data: blob:",
    "connect-src 'self' blob:",
    "object-src 'none'",
    "base-uri 'self'",
    "form-action 'self'",
    "frame-ancestors 'none'",
  ].join("; ");

  const devHeaders = {
    "Content-Security-Policy": devCsp,
    "X-Frame-Options": "DENY",
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "strict-origin-when-cross-origin",
  };

  const proxy = {
    "/api": {
      target: apiTarget,
      changeOrigin: true,
      // Server-sent events must stream, not buffer.
      ws: true,
      rewrite: (path) => path.replace(/^\/api/, ""),
    },
  };

  return {
    server: { headers: devHeaders, proxy },
    preview: { headers: devHeaders, proxy },
  };
});
