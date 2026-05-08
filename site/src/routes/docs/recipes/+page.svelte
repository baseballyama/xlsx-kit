<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>Recipes — xlsxify</title>
</svelte:head>

<div class="content">
  <header class="page-head">
    <p class="eyebrow">Cookbook</p>
    <h1>Recipes</h1>
    <p class="lede">
      Working code for the things people actually want to do — open a workbook, build one
      from scratch, style cells, add a chart, stream millions of rows. Every snippet on
      this page is a real <code>.ts</code> file in the repo, type-checked against the
      live <code>xlsxify</code> on every build, so what you see compiles.
    </p>
    <p class="meta">
      {data.groups.flatMap((g) => g.recipes).length} recipes across {data.groups.length}
      categories. Looking up a specific function? Jump to the
      <a href="{base}/api">API reference</a>.
    </p>
  </header>

  <nav class="toc" aria-label="Recipes index">
    <h4>Recipes</h4>
    {#each data.groups as group (group.title)}
      <div class="toc-group">
        <span class="toc-group-label">{group.title}</span>
        <ul>
          {#each group.recipes as r (r.slug)}
            <li>
              <a href="#{r.slug}">{r.title}</a>
            </li>
          {/each}
        </ul>
      </div>
    {/each}
  </nav>

  <div class="recipes">
    {#each data.groups as group (group.title)}
      <section class="group" id={'group-' + group.title.toLowerCase().replace(/\s+/g, '-')}>
        <h2>{group.title}</h2>
        {#each group.recipes as r (r.slug)}
          <article class="recipe" id={r.slug}>
            <header>
              <h3>
                <a href="#{r.slug}" class="hash" aria-label="Permalink">#</a>
                {r.title}
              </h3>
              <p class="teaser">{r.teaser}</p>
            </header>
            <CodeBlock html={r.html} title={r.path} />
            {#if r.notes?.length}
              <ul class="notes">
                {#each r.notes as n (n)}
                  <li>{n}</li>
                {/each}
              </ul>
            {/if}
            {#if r.relatedApi?.length}
              <p class="related">
                <span class="related-label">Related API:</span>
                {#each r.relatedApi as name, i (name)}
                  <code>{name}</code>{#if i < r.relatedApi.length - 1}, {/if}
                {/each}
              </p>
            {/if}
          </article>
        {/each}
      </section>
    {/each}
  </div>
</div>

<style>
  .content {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 260px;
    grid-template-areas:
      'head head'
      'recipes toc';
    column-gap: 2rem;
    max-width: 1280px;
    padding: 2rem 1.5rem 5rem;
  }

  .page-head {
    grid-area: head;
    margin-bottom: 1.5rem;
  }

  .recipes {
    grid-area: recipes;
    min-width: 0;
  }

  .toc {
    grid-area: toc;
    position: sticky;
    top: calc(var(--header-h) + 1.5rem);
    align-self: start;
    max-height: calc(100vh - var(--header-h) - 3rem);
    overflow-y: auto;
    font-size: 13px;
    padding-left: 1rem;
    border-left: 1px solid var(--border);
  }

  .eyebrow {
    color: var(--accent);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-size: 13px;
    margin: 0 0 0.4rem;
  }

  h1 {
    margin: 0 0 0.4rem;
  }

  .lede {
    color: var(--fg-soft);
    margin: 0;
    max-width: 720px;
  }

  .meta {
    color: var(--fg-muted);
    font-size: 0.9rem;
    margin: 0.8rem 0 0;
  }

  .toc h4 {
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--fg-muted);
    margin: 0 0 0.6rem;
    border: none;
  }

  .toc-group {
    margin-bottom: 1rem;
  }

  .toc-group-label {
    display: block;
    font-weight: 600;
    color: var(--fg);
    padding: 0.2rem 0.4rem;
    margin-bottom: 0.2rem;
    font-size: 13px;
  }

  .toc ul {
    list-style: none;
    padding: 0 0 0 0.7rem;
    margin: 0;
    border-left: 1px solid var(--border);
  }

  .toc li a {
    display: block;
    padding: 0.18rem 0.4rem;
    color: var(--fg-soft);
    line-height: 1.35;
    font-size: 13px;
    border-radius: 4px;
  }

  .toc li a:hover {
    color: var(--fg);
    background: var(--bg-soft);
    text-decoration: none;
  }

  h2 {
    border-bottom: 2px solid var(--accent);
    padding-bottom: 0.4rem;
    margin-top: 2.5rem;
  }

  .recipe {
    border-top: 1px solid var(--border);
    padding: 1.6rem 0 0.8rem;
    scroll-margin-top: calc(var(--header-h) + 1rem);
  }

  .recipe h3 {
    margin: 0 0 0.3rem;
    font-size: 1.2rem;
  }

  .hash {
    color: var(--fg-muted);
    font-weight: 400;
    margin-right: 0.3rem;
    text-decoration: none;
    visibility: hidden;
  }

  .recipe h3:hover .hash {
    visibility: visible;
  }

  .teaser {
    color: var(--fg-soft);
    margin: 0.2rem 0 0.6rem;
  }

  .notes {
    margin: 0.8rem 0 0.4rem;
    padding-left: 1.2rem;
    color: var(--fg-soft);
    font-size: 0.92rem;
  }

  .notes li {
    margin-bottom: 0.3rem;
  }

  .related {
    margin: 0.8rem 0 0;
    color: var(--fg-muted);
    font-size: 0.85rem;
  }

  .related-label {
    color: var(--fg-muted);
    margin-right: 0.4rem;
  }

  .related code {
    margin-right: 0.15rem;
  }

  @media (max-width: 1000px) {
    .content {
      grid-template-columns: 1fr;
      grid-template-areas:
        'head'
        'toc'
        'recipes';
    }
    .toc {
      position: static;
      max-height: none;
      border-left: none;
      padding-left: 0;
      border-top: 1px solid var(--border);
      border-bottom: 1px solid var(--border);
      padding: 0.8rem 0;
    }
  }
</style>
