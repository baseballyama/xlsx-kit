<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>openxml-js — TypeScript port of openpyxl</title>
</svelte:head>

<section class="hero">
  <div class="hero-inner">
    <p class="eyebrow">A TypeScript port of openpyxl</p>
    <h1>Read and write Excel <code class="hero-code">.xlsx</code> from Node and the browser.</h1>
    <p class="lede">
      Full workbook model — values, formulas, styles, charts, drawings, pivots, VBA — plus a
      streaming writer that pushes 10M rows under a 100&nbsp;MB heap. No Python, no Excel, no
      runtime native modules.
    </p>
    <div class="cta">
      <a href="{base}/docs/getting-started" class="btn primary">Get started →</a>
      <a href="{base}/docs/recipes" class="btn">Recipes</a>
      <a href="{base}/api" class="btn">API reference</a>
      <a href="https://github.com/baseballyama/openxml-js" class="btn">GitHub</a>
    </div>
    <p class="install"><code>pnpm add openxml-js</code></p>
  </div>
</section>

<section class="features">
  <div class="features-inner">
    <article>
      <h3>Round-trips real workbooks</h3>
      <p>
        Pivot tables, macro-enabled <code>.xlsm</code>, threaded comments, Power Query metadata,
        custom XML — anything we don't model is preserved byte-identical so Excel 365 still
        renders it.
      </p>
    </article>
    <article>
      <h3>Streaming, both ways</h3>
      <p>
        <code>createWriteOnlyWorkbook</code> deflates rows as they arrive.
        <code>loadWorkbookStream</code> walks a file once and yields each row. Browser-safe via
        <code>openxml-js/streaming</code>.
      </p>
    </article>
    <article>
      <h3>Charts &amp; drawings, modeled</h3>
      <p>
        16 legacy <code>c:</code> chart kinds plus 8 <code>cx:</code> chartex kinds (Sunburst,
        Treemap, Waterfall, Histogram, Pareto, Funnel, BoxWhisker, RegionMap). Images
        auto-detect format and dimensions.
      </p>
    </article>
    <article>
      <h3>Tiny &amp; tree-shakeable</h3>
      <p>
        <code>openxml-js</code> ≤ 120&nbsp;KB brotli (currently ~78&nbsp;KB).
        <code>openxml-js/streaming</code> ≤ 80&nbsp;KB brotli (~47&nbsp;KB). All exports are
        side-effect-free.
      </p>
    </article>
  </div>
</section>

<section class="examples">
  <div class="examples-inner">
    <h2>Two snippets to get the shape</h2>
    <p class="lede">
      Both files below live under <code>site/src/lib/examples/</code> and are type-checked by
      <code>svelte-check</code> against the real library on every build — if an API renames,
      the docs build fails.
    </p>
    {#each data.hero as ex (ex.key)}
      <h3>{ex.title}</h3>
      <p>{ex.description}</p>
      <CodeBlock html={ex.html} title={ex.path} />
    {/each}
    <p class="more">
      More examples in <a href="{base}/docs/getting-started">Getting started</a> and
      <a href="{base}/docs/streaming">Streaming</a>.
    </p>
  </div>
</section>

<style>
  .hero {
    padding: 4rem 1.25rem 3rem;
    border-bottom: 1px solid var(--border);
    background:
      radial-gradient(circle at 25% 0%, rgba(255, 62, 0, 0.16), transparent 50%),
      radial-gradient(circle at 80% 30%, rgba(80, 90, 220, 0.12), transparent 60%);
  }

  .hero-inner {
    max-width: 880px;
    margin: 0 auto;
    text-align: center;
  }

  .eyebrow {
    color: var(--accent);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-size: 13px;
    margin: 0 0 0.6rem;
  }

  .hero h1 {
    font-size: clamp(2rem, 4.5vw, 3.25rem);
    margin: 0 0 1rem;
    line-height: 1.15;
  }

  .hero-code {
    background: var(--bg-soft);
    border: 1px solid var(--border);
    padding: 0.05em 0.3em;
    font-size: 0.85em;
    color: var(--accent);
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.075rem;
    max-width: 680px;
    margin: 0 auto 1.5rem;
  }

  .cta {
    display: flex;
    gap: 0.75rem;
    justify-content: center;
    flex-wrap: wrap;
    margin-bottom: 1.5rem;
  }

  .btn {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    padding: 0.65rem 1.1rem;
    border-radius: 999px;
    background: var(--bg-elev);
    border: 1px solid var(--border);
    color: var(--fg);
    font-weight: 500;
    font-size: 0.95rem;
  }

  .btn:hover {
    background: var(--bg-soft);
    text-decoration: none;
  }

  .btn.primary {
    background: var(--accent);
    border-color: var(--accent);
    color: white;
  }

  .btn.primary:hover {
    filter: brightness(1.1);
  }

  .install {
    font-family: var(--mono);
    color: var(--fg-muted);
    font-size: 0.95rem;
  }

  .install code {
    background: var(--bg-soft);
    padding: 0.4em 0.8em;
    font-size: 0.95em;
  }

  .features {
    padding: 4rem 1.25rem;
    background: var(--bg-elev);
    border-bottom: 1px solid var(--border);
  }

  .features-inner {
    max-width: 1100px;
    margin: 0 auto;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 1.5rem;
  }

  .features article {
    padding: 1.25rem;
    background: var(--bg-soft);
    border: 1px solid var(--border);
    border-radius: var(--radius);
  }

  .features h3 {
    margin-top: 0;
    font-size: 1.05rem;
  }

  .features p {
    color: var(--fg-soft);
    font-size: 0.93rem;
    margin: 0.4rem 0 0;
  }

  .examples {
    padding: 4rem 1.25rem 6rem;
  }

  .examples-inner {
    max-width: var(--max-content);
    margin: 0 auto;
  }

  .examples h2 {
    border-bottom: none;
    padding-bottom: 0;
    margin-top: 0;
  }

  .examples h3 {
    margin-top: 2.5rem;
  }

  .more {
    margin-top: 2rem;
    color: var(--fg-soft);
  }
</style>
