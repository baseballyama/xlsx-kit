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

  let bodyEl = $state<HTMLDivElement | undefined>();
  let status = $state<'idle' | 'copied' | 'failed'>('idle');
  let timer: ReturnType<typeof setTimeout> | undefined;

  async function copy() {
    const text = bodyEl?.querySelector('pre')?.innerText ?? '';
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      status = 'copied';
    } catch {
      status = 'failed';
    }
    clearTimeout(timer);
    timer = setTimeout(() => {
      status = 'idle';
    }, 1500);
  }
</script>

<figure class="code-block">
  <figcaption>
    {#if coord}
      <span class="coord">{coord}</span>
    {/if}
    {#if title}
      <span class="path">{title}</span>
      <span class="lang">.ts</span>
    {:else}
      <span class="path" aria-hidden="true"></span>
    {/if}
    <button
      type="button"
      class="copy"
      class:copied={status === 'copied'}
      class:failed={status === 'failed'}
      onclick={copy}
      aria-label="Copy code to clipboard"
    >
      {#if status === 'copied'}Copied{:else if status === 'failed'}Failed{:else}Copy{/if}
    </button>
  </figcaption>
  <div class="body" bind:this={bodyEl}>{@html html}</div>
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
    padding: 0.4rem 0.45rem 0.4rem 0.9rem;
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
    min-width: 0;
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

  .copy {
    font-family: var(--mono);
    font-size: 10.5px;
    font-weight: 500;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--fg-soft);
    background: transparent;
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0.28em 0.7em;
    cursor: pointer;
    transition:
      color 120ms ease,
      background 120ms ease,
      border-color 120ms ease;
  }

  .copy:hover {
    color: var(--fg);
    background: var(--bg-soft);
    border-color: var(--border-strong);
  }

  .copy.copied {
    color: var(--accent);
    border-color: var(--accent-soft);
    background: var(--accent-soft);
  }

  .copy.failed {
    color: var(--fg);
    border-color: var(--border-strong);
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
