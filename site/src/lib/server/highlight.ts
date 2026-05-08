// Server-only Shiki helper. Used by +page.server.ts loaders to highlight
// raw source files at prerender time, so the client never ships Shiki.
import { createHighlighter, type Highlighter } from 'shiki';

const THEME = 'github-dark';
const LANGS = ['ts', 'tsx', 'js', 'json', 'sh', 'bash', 'xml', 'svelte', 'html'] as const;

let highlighterPromise: Promise<Highlighter> | null = null;

function getHighlighter(): Promise<Highlighter> {
  highlighterPromise ??= createHighlighter({ themes: [THEME], langs: [...LANGS] });
  return highlighterPromise;
}

export async function highlight(code: string, lang: string = 'ts'): Promise<string> {
  const safe = (LANGS as readonly string[]).includes(lang) ? lang : 'text';
  const h = await getHighlighter();
  return h.codeToHtml(code, { lang: safe, theme: THEME });
}
