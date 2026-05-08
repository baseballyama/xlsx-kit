import { loadApiSections } from '$lib/api/parse';
import type { ApiSectionSummary } from '$lib/api/types';
import type { LayoutServerLoad } from './$types';

export const load: LayoutServerLoad = () => {
  const sections = loadApiSections();
  const summaries: ApiSectionSummary[] = sections.map((s) => ({
    id: s.id,
    title: s.title,
    description: s.description,
    itemCount: s.items.length,
  }));
  return { sections: summaries };
};
