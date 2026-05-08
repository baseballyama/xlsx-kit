<script lang="ts">
  type Props = {
    /** Pre-rendered Shiki HTML (the full <pre>...</pre>). */
    html: string;
    /** Optional file path / caption shown above the snippet. */
    title?: string;
  };

  const { html, title }: Props = $props();
</script>

<figure class="code-block">
  {#if title}
    <figcaption>
      <span class="dot"></span><span class="dot"></span><span class="dot"></span>
      <span class="title">{title}</span>
    </figcaption>
  {/if}
  <div class="body">{@html html}</div>
</figure>

<style>
  .code-block {
    margin: 1.25rem 0;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: var(--code-bg);
    overflow: hidden;
  }

  figcaption {
    display: flex;
    align-items: center;
    gap: 0.4rem;
    padding: 0.55rem 0.85rem;
    background: var(--bg-soft);
    border-bottom: 1px solid var(--border);
    font-family: var(--mono);
    font-size: 13px;
    color: var(--fg-muted);
  }

  .dot {
    width: 10px;
    height: 10px;
    border-radius: 50%;
    background: var(--border);
  }

  .title {
    margin-left: 0.5rem;
  }

  .body :global(pre) {
    margin: 0;
    border: none;
    border-radius: 0;
    background: transparent;
  }
</style>
