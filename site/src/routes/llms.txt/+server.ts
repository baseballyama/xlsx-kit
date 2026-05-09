// /llms.txt — short, well-formed index of every doc page.
// Format follows https://llmstxt.org/ proposal: H1 title, blockquote
// summary, then sections with bullet links + descriptions.

import { docSections } from '$lib/docs-nav';
import type { RequestHandler } from './$types';

export const prerender = true;

const HEADER = `# xlsx-kit

> Read and write Excel \`.xlsx\` workbooks from Node 22+ and modern browsers, with no Python or native runtime dependencies. Includes a streaming writer (10M rows in fixed memory) and a SAX-based row iterator for huge sheets.

This index points at the canonical documentation for the library. Every linked URL is also available as raw Markdown by appending \`.md\` (e.g. \`/docs/install.md\`). Append \`/llms-full.txt\` to this site to get every page concatenated into a single document.
`;

function buildBody(): string {
  // Relative paths so the index works under any base path (custom domain,
  // user page, or project page like /xlsx-kit/). `docs/install.md` from
  // /llms.txt resolves to <base>/docs/install.md regardless.
  const sections = docSections
    .map((section) => {
      const lines = section.links.map(
        (l) => `- [${l.title}](.${l.href}.md): ${l.description}`,
      );
      return `## ${section.title}\n\n${lines.join('\n')}`;
    })
    .join('\n\n');

  return `${HEADER}\n${sections}\n\n## Source\n\n- [GitHub repository](https://github.com/baseballyama/xlsx-kit)\n- [npm package](https://www.npmjs.com/package/xlsx-kit)\n`;
}

export const GET: RequestHandler = () => {
  return new Response(buildBody(), {
    headers: {
      'Content-Type': 'text/plain; charset=utf-8',
      'Cache-Control': 'public, max-age=300',
    },
  });
};
