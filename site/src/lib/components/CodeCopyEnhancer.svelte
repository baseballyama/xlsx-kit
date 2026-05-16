<script lang="ts">
  // Adds a floating "Copy" button to every <pre> rendered outside a CodeBlock
  // figure — i.e. every code fence inside an .svx / .md doc. CodeBlock has its
  // own copy button in the figcaption, so we skip those.
  import { afterNavigate } from '$app/navigation';
  import { onMount, tick } from 'svelte';

  const ENHANCED = 'data-copy-enhanced';

  async function enhance() {
    await tick();
    const pres = document.querySelectorAll<HTMLPreElement>(`pre:not([${ENHANCED}])`);
    for (const pre of pres) {
      if (pre.closest('.code-block')) continue;
      attach(pre);
    }
  }

  function attach(pre: HTMLPreElement) {
    pre.setAttribute(ENHANCED, '');

    // Wrap the <pre> so the absolutely-positioned button anchors to its corner
    // without depending on a containing block we don't control.
    const wrap = document.createElement('div');
    wrap.className = 'code-copy-wrap';
    pre.parentNode?.insertBefore(wrap, pre);
    wrap.appendChild(pre);

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'code-copy-btn';
    btn.setAttribute('aria-label', 'Copy code to clipboard');
    btn.textContent = 'Copy';

    let timer: ReturnType<typeof setTimeout> | undefined;
    btn.addEventListener('click', async () => {
      const text = pre.innerText;
      if (!text) return;
      try {
        await navigator.clipboard.writeText(text);
        btn.textContent = 'Copied';
        btn.classList.add('copied');
      } catch {
        btn.textContent = 'Failed';
        btn.classList.add('failed');
      }
      clearTimeout(timer);
      timer = setTimeout(() => {
        btn.textContent = 'Copy';
        btn.classList.remove('copied', 'failed');
      }, 1500);
    });

    wrap.appendChild(btn);
  }

  onMount(() => void enhance());
  afterNavigate(() => void enhance());
</script>

<style>
  :global(.code-copy-wrap) {
    position: relative;
  }

  :global(.code-copy-btn) {
    position: absolute;
    top: 0.5rem;
    right: 0.5rem;
    font-family: var(--mono);
    font-size: 10.5px;
    font-weight: 500;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--fg-soft);
    background: color-mix(in oklab, var(--bg-paper) 92%, transparent);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0.28em 0.7em;
    cursor: pointer;
    opacity: 0;
    transition:
      opacity 120ms ease,
      color 120ms ease,
      background 120ms ease,
      border-color 120ms ease;
  }

  :global(.code-copy-wrap:hover .code-copy-btn),
  :global(.code-copy-btn:focus-visible),
  :global(.code-copy-btn.copied),
  :global(.code-copy-btn.failed) {
    opacity: 1;
  }

  :global(.code-copy-btn:hover) {
    color: var(--fg);
    background: var(--bg-soft);
    border-color: var(--border-strong);
  }

  :global(.code-copy-btn.copied) {
    color: var(--accent);
    border-color: var(--accent-soft);
    background: var(--accent-soft);
  }
</style>
