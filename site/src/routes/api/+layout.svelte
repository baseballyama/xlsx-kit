<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import type { LayoutProps } from './$types';

  const { data, children }: LayoutProps = $props();

  const currentSection = $derived(page.params.section ?? null);
</script>

<div class="layout">
  <aside class="sidebar">
    <h4>API reference</h4>
    <ul>
      <li>
        <a href="{base}/api" class:active={currentSection === null}>Overview</a>
      </li>
    </ul>
    <ul>
      {#each data.sections as section (section.id)}
        <li>
          <a
            href="{base}/api/{section.id}"
            class:active={currentSection === section.id}
          >
            <span class="title">{section.title}</span>
            <span class="count">{section.itemCount}</span>
          </a>
        </li>
      {/each}
    </ul>
  </aside>
  <div class="main">
    {@render children?.()}
  </div>
</div>

<style>
  .layout {
    display: flex;
    align-items: stretch;
    max-width: var(--max-wide);
    margin: 0 auto;
  }

  .sidebar {
    width: 268px;
    flex: 0 0 268px;
    padding: 2rem 0.75rem 4rem 1.5rem;
    height: calc(100vh - var(--header-h));
    position: sticky;
    top: var(--header-h);
    overflow-y: auto;
  }

  h4 {
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    color: var(--fg-soft);
    margin: 0 0 0.6rem 0.5rem;
    border: none;
    padding: 0;
    font-variation-settings: normal;
  }

  ul {
    list-style: none;
    padding: 0;
    margin: 0 0 1rem 0;
  }

  li a {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.5rem;
    padding: 0.4rem 0.6rem;
    color: var(--fg-soft);
    border-radius: var(--radius-sm);
    font-size: 0.92rem;
    line-height: 1.35;
    border-left: 2px solid transparent;
  }

  li a:hover {
    background: var(--bg-soft);
    color: var(--fg);
    text-decoration: none;
  }

  li a.active {
    background: var(--accent-soft);
    color: var(--fg);
    border-left-color: var(--accent);
  }

  .title {
    flex: 1;
    min-width: 0;
  }

  .count {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    background: var(--bg-paper);
    border: 1px solid var(--border);
    padding: 0.05em 0.4em;
    border-radius: 3px;
    font-variant-numeric: tabular-nums;
    letter-spacing: 0.04em;
  }

  .main {
    flex: 1;
    min-width: 0;
    padding: 0 1rem;
  }

  @media (max-width: 800px) {
    .layout {
      flex-direction: column;
    }
    .sidebar {
      width: 100%;
      flex: none;
      position: static;
      height: auto;
      border-bottom: 1px solid var(--border);
      padding: 1rem;
    }
  }
</style>
