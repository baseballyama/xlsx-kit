<script lang="ts">
  type Props = {
    /** Pre-rendered Shiki HTML (the full <pre>...</pre>). */
    html: string;
    /** Optional file path / caption shown above the snippet. */
    title?: string;
    /** Optional cell-coordinate label (defaults to next sequential coord). */
    coord?: string;
  };

  const { html, title, coord }: Props = $props();
</script>

<figure class="code-block">
  {#if title}
    <figcaption>
      {#if coord}
        <span class="coord">{coord}</span>
      {/if}
      <span class="path">{title}</span>
      <span class="lang">.ts</span>
    </figcaption>
  {/if}
  <div class="body">{@html html}</div>
</figure>

<style>
  .code-block {
    margin: 1.4rem 0;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: var(--code-bg);
    overflow: hidden;
    box-shadow: 0 1px 0 0 var(--border-strong) inset;
  }

  figcaption {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.55rem 0.9rem;
    background: var(--bg-paper);
    border-bottom: 1px solid var(--border);
    font-family: var(--mono);
    font-size: 12px;
    color: var(--fg-soft);
    letter-spacing: 0.02em;
  }

  .coord {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 2.6ch;
    padding: 0.1em 0.4em;
    font-size: 11px;
    font-weight: 500;
    color: var(--accent);
    background: var(--accent-soft);
    border: 1px solid var(--accent-soft);
    border-radius: 3px;
    letter-spacing: 0.04em;
  }

  .path {
    color: var(--fg);
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .lang {
    color: var(--fg-muted);
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    font-weight: 500;
  }

  .body :global(pre) {
    margin: 0;
    border: none;
    border-radius: 0;
    background: transparent !important;
    padding: 1rem 1.1rem;
  }

  .body :global(pre code) {
    font-family: var(--mono);
  }
</style>
