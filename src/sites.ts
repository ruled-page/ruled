/**
 * Where the example sites are, for the links between them. In development
 * they sit side by side on one server; published, each is its own Cloudflare
 * Pages site and the build passes their addresses (see scripts/build-sites.mjs).
 * The Angular app has no development page here, so it is only linked once published.
 */
export const SITES = {
  demo: (import.meta.env.VITE_DEMO_URL as string | undefined) ?? import.meta.env.BASE_URL,
  portal: (import.meta.env.VITE_PORTAL_URL as string | undefined) ?? `${import.meta.env.BASE_URL}embedded.html`,
  angular: import.meta.env.VITE_ANGULAR_URL as string | undefined,
};
