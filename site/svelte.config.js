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
// or custom domain (=''), or on a project page (e.g. '/openxml-js'). Set
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
      'openxml-js/cell': '../src/cell/index.ts',
      'openxml-js/chart': '../src/chart/index.ts',
      'openxml-js/chartsheet': '../src/chartsheet/index.ts',
      'openxml-js/drawing': '../src/drawing/index.ts',
      'openxml-js/io': '../src/io/index.ts',
      'openxml-js/node': '../src/node.ts',
      'openxml-js/packaging': '../src/packaging/index.ts',
      'openxml-js/schema': '../src/schema/index.ts',
      'openxml-js/streaming': '../src/streaming/index.ts',
      'openxml-js/styles': '../src/styles/index.ts',
      'openxml-js/utils': '../src/utils/index.ts',
      'openxml-js/workbook': '../src/workbook/index.ts',
      'openxml-js/worksheet': '../src/worksheet/index.ts',
      'openxml-js/xml': '../src/xml/index.ts',
      'openxml-js/zip': '../src/zip/index.ts',
    },
  },
};

export default config;
