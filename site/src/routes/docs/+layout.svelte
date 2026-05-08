<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import Sidebar from '$lib/components/Sidebar.svelte';
  import { allDocLinks } from '$lib/docs-nav';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();

  const mdHref = $derived(`${page.url.pathname.replace(/\/$/, '')}.md`);
  const currentRoute = $derived(page.url.pathname.replace(new RegExp(`^${base}`), '') || '/');
  const knownDoc = $derived(allDocLinks.some((l) => l.href === currentRoute));
</script>

<div class="layout">
  <Sidebar />
  <div class="main">
    <article class="doc-content">
      {@render children?.()}
    </article>
    {#if knownDoc}
      <div class="md-link">
        <a href={mdHref}>View raw Markdown</a>
        <span class="muted">
          (LLMs and tools can fetch this URL or
          <a href="{base}/llms.txt">/llms.txt</a> for the full index)
        </span>
      </div>
    {/if}
  </div>
</div>

<style>
  .layout {
    display: flex;
    align-items: stretch;
    max-width: 1300px;
    margin: 0 auto;
  }

  .main {
    flex: 1;
    min-width: 0;
    padding: 0 1rem;
  }

  .doc-content {
    max-width: var(--max-content);
    padding: 2rem 1.5rem 3rem;
  }

  .doc-content :global(h2) {
    margin-top: 2.5rem;
  }

  .md-link {
    max-width: var(--max-content);
    padding: 1.5rem 1.5rem 3rem;
    border-top: 1px solid var(--border);
    margin-top: 2rem;
    font-size: 0.92rem;
  }

  .md-link a {
    font-weight: 600;
    margin-right: 0.5rem;
  }

  .muted {
    color: var(--fg-muted);
  }

  @media (max-width: 800px) {
    .layout {
      flex-direction: column;
    }
  }
</style>
