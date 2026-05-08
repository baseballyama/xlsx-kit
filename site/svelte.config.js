import adapter from '@sveltejs/adapter-static';
import { vitePreprocess } from '@sveltejs/vite-plugin-svelte';
import { mdsvex, escapeSvelte } from 'mdsvex';
import { createHighlighter } from 'shiki';

const theme = 'github-dark';
const langs = ['ts', 'tsx', 'js', 'json', 'sh', 'bash', 'xml', 'svelte', 'html'];

const highlighter = await createHighlighter({ themes: [theme], langs });

/** @type {import('mdsvex').MdsvexOptions} */
const mdsvexOptions = {
  extensions: ['.svx', '.md'],
  highlight: {
    highlighter: async (code, lang = 'text') => {
      const safeLang = langs.includes(lang) ? lang : 'text';
      const html = escapeSvelte(highlighter.codeToHtml(code, { lang: safeLang, theme }));
      return `{@html \`${html}\`}`;
    },
  },
};

// BASE_PATH lets the same build run locally (=''), on a GitHub user page
// or custom domain (=''), or on a project page (e.g. '/xlsxlite'). Set
// it in CI for GitHub Actions deploys.
const basePath = process.env.BASE_PATH ?? '';

/** @type {import('@sveltejs/kit').Config} */
const config = {
  extensions: ['.svelte', '.svx', '.md'],
  preprocess: [vitePreprocess(), mdsvex(mdsvexOptions)],
  kit: {
    adapter: adapter({ fallback: '404.html' }),
    prerender: { entries: ['*'] },
    paths: { base: basePath, relative: true },
    alias: {
      'xlsxlite/cell': '../src/cell/index.ts',
      'xlsxlite/chart': '../src/chart/index.ts',
      'xlsxlite/chartsheet': '../src/chartsheet/index.ts',
      'xlsxlite/drawing': '../src/drawing/index.ts',
      'xlsxlite/io': '../src/io/index.ts',
      'xlsxlite/node': '../src/node.ts',
      'xlsxlite/packaging': '../src/packaging/index.ts',
      'xlsxlite/schema': '../src/schema/index.ts',
      'xlsxlite/streaming': '../src/streaming/index.ts',
      'xlsxlite/styles': '../src/styles/index.ts',
      'xlsxlite/utils': '../src/utils/index.ts',
      'xlsxlite/workbook': '../src/workbook/index.ts',
      'xlsxlite/worksheet': '../src/worksheet/index.ts',
      'xlsxlite/xml': '../src/xml/index.ts',
      'xlsxlite/zip': '../src/zip/index.ts',
    },
  },
};

export default config;
