import { error } from '@sveltejs/kit';
import { loadApiSection, loadApiSections } from '$lib/api/parse';
import { highlight } from '$lib/server/highlight';
import type { ApiItem } from '$lib/api/types';
import type { PageServerLoad, EntryGenerator } from './$types';

export const prerender = true;

export const entries: EntryGenerator = () =>
  loadApiSections().map((s) => ({ section: s.id }));

export type RenderedItem = ApiItem & { signatureHtml: string };

export type RenderedSubgroup = {
  label: string;
  id: string;
  sourceFile: string;
  items: RenderedItem[];
};

export const load: PageServerLoad = async ({ params }) => {
  const section = loadApiSection(params.section);
  if (!section) throw error(404, `Unknown API section: ${params.section}`);

  const subgroups: RenderedSubgroup[] = await Promise.all(
    section.subgroups.map(async (group) => ({
      label: group.label,
      id: group.id,
      sourceFile: group.sourceFile,
      items: await Promise.all(
        group.items.map(async (item) => ({
          ...item,
          signatureHtml: await highlight(item.signature, 'ts'),
        })),
      ),
    })),
  );

  return {
    section: {
      id: section.id,
      title: section.title,
      description: section.description,
      itemCount: section.items.length,
    },
    subgroups,
  };
};
