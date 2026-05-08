// Pre-highlight every example file once so each docs page can pick which
// snippets to render without paying the Shiki cost per page.

import { examples, type Example, type ExampleKey } from '$lib/examples';
import { highlight } from '$lib/server/highlight';
import type { LayoutServerLoad } from './$types';

export type HighlightedExample = Example & { html: string };

export const load: LayoutServerLoad = async () => {
  const keys = Object.keys(examples) as ExampleKey[];
  const entries = await Promise.all(
    keys.map(async (key) => {
      const ex = examples[key];
      const html = await highlight(ex.source, 'ts');
      return [key, { ...ex, html }] as const;
    }),
  );
  const out = Object.fromEntries(entries) as Record<ExampleKey, HighlightedExample>;
  return { examples: out };
};
