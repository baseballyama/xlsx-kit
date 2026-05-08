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
// or custom domain (=''), or on a project page (e.g. '/xlsxify'). Set
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
      'xlsxify/cell': '../src/cell/index.ts',
      'xlsxify/chart': '../src/chart/index.ts',
      'xlsxify/chartsheet': '../src/chartsheet/index.ts',
      'xlsxify/drawing': '../src/drawing/index.ts',
      'xlsxify/io': '../src/io/index.ts',
      'xlsxify/node': '../src/node.ts',
      'xlsxify/packaging': '../src/packaging/index.ts',
      'xlsxify/schema': '../src/schema/index.ts',
      'xlsxify/streaming': '../src/streaming/index.ts',
      'xlsxify/styles': '../src/styles/index.ts',
      'xlsxify/utils': '../src/utils/index.ts',
      'xlsxify/workbook': '../src/workbook/index.ts',
      'xlsxify/worksheet': '../src/worksheet/index.ts',
      'xlsxify/xml': '../src/xml/index.ts',
      'xlsxify/zip': '../src/zip/index.ts',
    },
  },
};

export default config;
