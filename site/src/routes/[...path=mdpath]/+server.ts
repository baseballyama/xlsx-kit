// Catch-all that turns `<docs page>.md` URLs into raw Markdown. Lets LLMs
// (and humans) fetch the source of any docs page without scraping HTML.

import { error } from '@sveltejs/kit';
import { allDocLinks } from '$lib/docs-nav';
import { getMarkdownDoc } from '$lib/server/markdown';
import type { RequestHandler, EntryGenerator } from './$types';

export const prerender = true;

export const entries: EntryGenerator = () =>
  allDocLinks.map(({ href }) => ({ path: `${href.slice(1)}.md` }));

export const GET: RequestHandler = ({ params }) => {
  const route = '/' + params.path.replace(/\.md$/, '');
  const doc = getMarkdownDoc(route);
  if (!doc) throw error(404, `No Markdown source for ${route}`);

  return new Response(doc.markdown, {
    headers: {
      'Content-Type': 'text/markdown; charset=utf-8',
      'Cache-Control': 'public, max-age=300',
    },
  });
};
