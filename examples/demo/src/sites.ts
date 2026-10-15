/**
 * Where the example sites are, for the links between them. In development
 * each runs on its own server (npm run dev, npm run dev:portal); published,
 * each is its own Cloudflare Pages site and the build passes their addresses
 * (see scripts/build-sites.mjs). The Angular app has no development server
 * here, so it is only linked once published.
 */
export const SITES = {
  demo: (import.meta.env.VITE_DEMO_URL as string | undefined) ?? 'http://localhost:5180/',
  portal: (import.meta.env.VITE_PORTAL_URL as string | undefined) ?? 'http://localhost:5181/',
  angular: import.meta.env.VITE_ANGULAR_URL as string | undefined,
  docs: (import.meta.env.VITE_DOCS_URL as string | undefined) ?? 'https://ruled.page/',
};
