// /llms-full.txt — every docs page concatenated into one document. Lets a
// model ingest the whole site in a single fetch.

import { listMarkdownDocs } from '$lib/server/markdown';
import type { RequestHandler } from './$types';

export const prerender = true;

const PREAMBLE = `# xlsx-craft — full documentation

This file is the concatenation of every page on https://xlsx-craft's docs site, intended for LLM ingestion. Page boundaries are marked with H1 headings prefixed by the source path.

Source repo: https://github.com/baseballyama/xlsx-craft
`;

export const GET: RequestHandler = () => {
  const docs = listMarkdownDocs();
  const parts = docs.map((doc) => {
    const title = doc.link?.title ?? doc.route;
    return [
      `\n\n---\n`,
      `<!-- Page: ${doc.route} -->`,
      `# ${title}`,
      doc.link?.description ? `> ${doc.link.description}\n` : '',
      doc.markdown,
    ].join('\n');
  });

  const body = `${PREAMBLE}\n${parts.join('\n')}`;

  return new Response(body, {
    headers: {
      'Content-Type': 'text/plain; charset=utf-8',
      'Cache-Control': 'public, max-age=300',
    },
  });
};
