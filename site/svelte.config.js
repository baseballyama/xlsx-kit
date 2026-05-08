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
// or custom domain (=''), or on a project page (e.g. '/xlsx-craft'). Set
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
      'xlsx-craft/cell': '../src/cell/index.ts',
      'xlsx-craft/chart': '../src/chart/index.ts',
      'xlsx-craft/chartsheet': '../src/chartsheet/index.ts',
      'xlsx-craft/drawing': '../src/drawing/index.ts',
      'xlsx-craft/io': '../src/io/index.ts',
      'xlsx-craft/node': '../src/node.ts',
      'xlsx-craft/packaging': '../src/packaging/index.ts',
      'xlsx-craft/schema': '../src/schema/index.ts',
      'xlsx-craft/streaming': '../src/streaming/index.ts',
      'xlsx-craft/styles': '../src/styles/index.ts',
      'xlsx-craft/utils': '../src/utils/index.ts',
      'xlsx-craft/workbook': '../src/workbook/index.ts',
      'xlsx-craft/worksheet': '../src/worksheet/index.ts',
      'xlsx-craft/xml': '../src/xml/index.ts',
      'xlsx-craft/zip': '../src/zip/index.ts',
    },
  },
};

export default config;
