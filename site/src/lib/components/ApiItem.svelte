<script lang="ts">
  import type { RenderedItem } from '../../routes/api/[section]/+page.server';

  type Props = {
    item: RenderedItem;
    /** Anchor id for the heading; usually `${subgroup}-${item.name}` so
     *  that same-name exports from different files don't collide. */
    anchorId: string;
  };

  const { item, anchorId }: Props = $props();

  const kindLabel: Record<string, string> = {
    function: 'function',
    class: 'class',
    interface: 'interface',
    type: 'type',
    variable: 'const',
  };
</script>

<section id={anchorId} class="item">
  <header>
    <h3>
      <a href="#{anchorId}" class="hash" aria-label="Permalink">#</a>
      <span class="name">{item.name}</span>
      <span class="kind kind-{item.kind}">{kindLabel[item.kind]}</span>
    </h3>
    <a href={item.sourceUrl} target="_blank" rel="noopener" class="source">
      {item.sourceFile}:{item.sourceLine}
    </a>
  </header>

  {#if item.description}
    <p class="description">{item.description}</p>
  {/if}

  <div class="signature">{@html item.signatureHtml}</div>

  {#if item.parameters?.length}
    <div class="params">
      <h4>Parameters</h4>
      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>Type</th>
            <th>Description</th>
          </tr>
        </thead>
        <tbody>
          {#each item.parameters as p (p.name)}
            <tr>
              <td>
                <code>{p.name}{p.optional ? '?' : ''}</code>
                {#if p.defaultValue}
                  <span class="default">= {p.defaultValue}</span>
                {/if}
              </td>
              <td><code>{p.type}</code></td>
              <td>{p.description ?? ''}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
  {/if}

  {#if item.returnType && item.kind === 'function'}
    <div class="returns">
      <h4>Returns</h4>
      <p><code>{item.returnType}</code>{item.returnDescription ? ` — ${item.returnDescription}` : ''}</p>
    </div>
  {/if}

</section>

<style>
  .item {
    border-top: 1px solid var(--border);
    padding: 2rem 0 1rem;
    scroll-margin-top: calc(var(--header-h) + 1rem);
  }

  header {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 1rem;
    flex-wrap: wrap;
    margin-bottom: 0.6rem;
  }

  h3 {
    margin: 0;
    font-size: 1.3rem;
    display: flex;
    align-items: baseline;
    gap: 0.5rem;
  }

  .hash {
    color: var(--fg-muted);
    font-weight: 400;
    text-decoration: none;
    visibility: hidden;
  }

  h3:hover .hash {
    visibility: visible;
  }

  .name {
    font-family: var(--mono);
  }

  .kind {
    font-family: var(--mono);
    font-size: 13px;
    padding: 0.1em 0.55em;
    border-radius: 999px;
    border: 1px solid var(--border);
    color: var(--fg-muted);
    background: var(--bg-soft);
    font-weight: 500;
    letter-spacing: 0.02em;
    align-self: center;
  }

  .kind-function {
    color: #a5d6a7;
    border-color: rgba(165, 214, 167, 0.4);
    background: rgba(165, 214, 167, 0.08);
  }
  .kind-interface {
    color: #90caf9;
    border-color: rgba(144, 202, 249, 0.4);
    background: rgba(144, 202, 249, 0.08);
  }
  .kind-type {
    color: #ce93d8;
    border-color: rgba(206, 147, 216, 0.4);
    background: rgba(206, 147, 216, 0.08);
  }
  .kind-class {
    color: #ffab91;
    border-color: rgba(255, 171, 145, 0.4);
    background: rgba(255, 171, 145, 0.08);
  }
  .kind-variable {
    color: #ffd54f;
    border-color: rgba(255, 213, 79, 0.4);
    background: rgba(255, 213, 79, 0.08);
  }

  .source {
    font-family: var(--mono);
    font-size: 13px;
    color: var(--fg-muted);
  }

  .source:hover {
    color: var(--accent);
  }

  .description {
    margin: 0 0 0.8rem;
    color: var(--fg-soft);
    line-height: 1.55;
    white-space: pre-wrap;
  }

  .signature :global(pre) {
    margin: 0.5rem 0 0.8rem;
    border-radius: var(--radius);
  }

  h4 {
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--fg-muted);
    margin: 1.2rem 0 0.4rem;
    border: none;
  }

  table {
    margin: 0.4rem 0;
    font-size: 14px;
  }

  td code,
  td .default {
    font-size: 13px;
  }

  .default {
    color: var(--fg-muted);
    margin-left: 0.4em;
  }
</style>
