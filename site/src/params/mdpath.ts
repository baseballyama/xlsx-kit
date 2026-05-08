import type { ParamMatcher } from '@sveltejs/kit';

// Only matches paths whose final segment ends in `.md`. Used by the
// catch-all route that serves raw Markdown for any docs page.
export const match: ParamMatcher = (value) => value.endsWith('.md');
