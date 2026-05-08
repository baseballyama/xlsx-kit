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
// or custom domain (=''), or on a project page (e.g. '/ooxml-js'). Set
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
      'ooxml-js/xlsx/cell': '../src/xlsx/cell/index.ts',
      'ooxml-js/xlsx/chart': '../src/xlsx/chart/index.ts',
      'ooxml-js/xlsx/chartsheet': '../src/xlsx/chartsheet/index.ts',
      'ooxml-js/xlsx/drawing': '../src/xlsx/drawing/index.ts',
      'ooxml-js/xlsx/io': '../src/xlsx/io/index.ts',
      'ooxml-js/xlsx/streaming': '../src/xlsx/streaming/index.ts',
      'ooxml-js/xlsx/styles': '../src/xlsx/styles/index.ts',
      'ooxml-js/xlsx/workbook': '../src/xlsx/workbook/index.ts',
      'ooxml-js/xlsx/worksheet': '../src/xlsx/worksheet/index.ts',
      'ooxml-js/io': '../src/io/index.ts',
      'ooxml-js/node': '../src/node.ts',
      'ooxml-js/packaging': '../src/packaging/index.ts',
      'ooxml-js/schema': '../src/schema/index.ts',
      'ooxml-js/utils': '../src/utils/index.ts',
      'ooxml-js/xml': '../src/xml/index.ts',
      'ooxml-js/zip': '../src/zip/index.ts',
    },
  },
};

export default config;
