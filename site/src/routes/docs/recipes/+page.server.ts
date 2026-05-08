// Pre-highlight every recipe snippet at build time so the client bundle
// doesn't ship Shiki. Each recipe also keeps its raw source so it can be
// served to LLMs via the .md endpoint.

import { recipeGroups, type Recipe } from '$lib/examples/recipes';
import { highlight } from '$lib/server/highlight';
import type { PageServerLoad } from './$types';

export type RenderedRecipe = Recipe & { html: string };
export type RenderedGroup = { title: string; recipes: RenderedRecipe[] };

export const load: PageServerLoad = async () => {
  const groups: RenderedGroup[] = await Promise.all(
    recipeGroups.map(async (g) => ({
      title: g.title,
      recipes: await Promise.all(
        g.recipes.map(async (r) => ({
          ...r,
          html: await highlight(r.source, 'ts'),
        })),
      ),
    })),
  );
  return { groups };
};
