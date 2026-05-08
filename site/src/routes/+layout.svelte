<script lang="ts">
  import '../app.css';
  import { base } from '$app/paths';
  import SiteHeader from '$lib/components/SiteHeader.svelte';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();
</script>

<SiteHeader />

<main>
  {@render children?.()}
</main>

<footer>
  <div class="footer-inner">
    <p>
      <a href="https://github.com/baseballyama/xlsx-craft">xlsx-craft</a> · MIT licensed · A
      TypeScript port of <a href="https://openpyxl.readthedocs.io/">openpyxl</a>.
    </p>
    <p class="muted">
      Looking for an LLM-friendly index? See <a href="{base}/llms.txt">/llms.txt</a> or
      <a href="{base}/llms-full.txt">/llms-full.txt</a>. Append <code>.md</code> to any
      docs URL for the raw Markdown source.
    </p>
  </div>
</footer>

<style>
  main {
    min-height: calc(100vh - var(--header-h));
  }

  footer {
    border-top: 1px solid var(--border);
    padding: 2rem 1.25rem;
    background: var(--bg-elev);
  }

  .footer-inner {
    max-width: 1200px;
    margin: 0 auto;
  }

  footer p {
    margin: 0.25rem 0;
    color: var(--fg-soft);
    font-size: 0.9rem;
  }

  footer .muted {
    color: var(--fg-muted);
  }
</style>
