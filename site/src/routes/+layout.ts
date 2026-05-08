// Static export: prerender every route, including dynamic ones reachable
// from the link graph. SvelteKit also prerenders +server.ts endpoints
// reachable from prerendered pages (or listed in `entries`).
export const prerender = true;
