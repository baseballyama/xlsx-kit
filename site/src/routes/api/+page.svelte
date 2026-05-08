<script lang="ts">
  import { base } from '$app/paths';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>API reference — ooxml-js</title>
</svelte:head>

<div class="content">
  <h1>API reference</h1>
  <p class="lede">
    Every public export of <code>ooxml-js</code>, <code>ooxml-js/streaming</code>, and
    <code>ooxml-js/node</code>, organized by section. Generated from the source via typedoc
    and rendered inline so types stay in lockstep with the package.
  </p>

  <div class="grid">
    {#each data.sections as section (section.id)}
      <a href="{base}/api/{section.id}" class="card">
        <div class="card-head">
          <h3>{section.title}</h3>
          <span class="count">{section.itemCount}</span>
        </div>
        <p>{section.description}</p>
      </a>
    {/each}
  </div>
</div>

<style>
  .content {
    max-width: 1100px;
    padding: 2rem 1.5rem 5rem;
  }

  h1 {
    margin: 0 0 0.6rem;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.05rem;
    margin: 0 0 2rem;
    max-width: 720px;
  }

  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    gap: 1rem;
  }

  .card {
    display: block;
    padding: 1.1rem 1.2rem;
    background: var(--bg-elev);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    color: var(--fg);
    transition:
      border-color 120ms,
      transform 120ms;
  }

  .card:hover {
    border-color: var(--accent);
    text-decoration: none;
    transform: translateY(-1px);
  }

  .card-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.5rem;
    margin-bottom: 0.4rem;
  }

  .card h3 {
    margin: 0;
    font-size: 1rem;
    line-height: 1.3;
  }

  .card p {
    margin: 0;
    color: var(--fg-soft);
    font-size: 0.9rem;
    line-height: 1.45;
  }

  .count {
    font-size: 13px;
    color: var(--fg-muted);
    background: var(--bg-soft);
    border: 1px solid var(--border);
    padding: 0.1em 0.55em;
    border-radius: 999px;
    font-variant-numeric: tabular-nums;
  }
</style>
